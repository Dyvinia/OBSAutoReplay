import obspython as obs
import os
import re
import win32gui
import win32process
import psutil
from windows_toasts import Toast, WindowsToaster, ToastDuration

def script_description():
    return """Auto start Replays Buffer when Fullscreen Game Capture displays output, as well as showing windows toasts when starting, stopping, and saving.

Author: Dyvinia"""

toaster = WindowsToaster('OBS Replay')

def script_load(settings):
    obs.obs_frontend_add_event_callback(obs_frontend_callback)
    obs.timer_add(auto_replay_buffer, 10000)

def script_unload():
    obs.timer_remove(auto_replay_buffer)
    toaster.clear_toasts()

def obs_frontend_callback(event):
    if event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
        path = move_recording()
        newToast = Toast()
        newToast.text_fields = ['Saved Replay', "Saved in " + path]
        newToast.duration = ToastDuration.Short
        toaster.clear_toasts()
        toaster.show_toast(newToast)
    elif event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_STARTED:
        newToast = Toast()
        newToast.text_fields = ['Started Replay Buffer']
        newToast.duration = ToastDuration.Short
        toaster.clear_toasts()
        toaster.show_toast(newToast)
    elif event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_STOPPED:
        newToast = Toast()
        newToast.text_fields = ['Stopped Replay Buffer']
        newToast.duration = ToastDuration.Short
        toaster.clear_toasts()
        toaster.show_toast(newToast)

def auto_replay_buffer():
    scene_as_source = obs.obs_frontend_get_current_scene()
    current_scene = obs.obs_scene_from_source(scene_as_source)
    scene_item = obs.obs_scene_find_source_recursive(current_scene, "Fullscreen Game Capture")
    source = obs.obs_sceneitem_get_source(scene_item)

    if (not obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) > 0):
        #print("Starting Replay Buffer")
        obs.obs_frontend_replay_buffer_start()
        
    elif (obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) == 0):
        #print("Stopping Replay Buffer")
        obs.obs_frontend_replay_buffer_stop()

    obs.obs_source_release(scene_as_source)

    #print(get_foreground_window())

def move_recording():
    replay_buffer = obs.obs_frontend_get_replay_buffer_output()
    cd = obs.calldata_create()
    ph = obs.obs_output_get_proc_handler(replay_buffer)
    obs.proc_handler_call(ph, "get_last_replay", cd)
    path = obs.calldata_string(cd, "path")
    obs.calldata_destroy(cd)
    obs.obs_output_release(replay_buffer)

    game = safe_for_path(get_foreground_window())
    new_path = os.path.dirname(path) + '/Replays/' + game + '/' + os.path.basename(path)

    print("Saving replay to: " + new_path)
    os.renames(path, new_path)

    return '/Replays/' + game + '/'

def get_foreground_window():
    try: 
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)

        return psutil.Process(pid).name().replace(".exe", "").replace('.', '').strip()
    except:
        return "Other"

def safe_for_path(s):
    fixed = re.sub(r'[/\\:*?"<>|.]', '', s).strip()
    if len(fixed) > 0:
        return fixed
    else:
        return "Other"