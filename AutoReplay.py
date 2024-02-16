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

def script_properties():
    props = obs.obs_properties_create()
    scene_for_clips = obs.obs_properties_add_list(props, "scene", "Scene For Clipping",
                                    obs.OBS_COMBO_TYPE_EDITABLE,
                                    obs.OBS_COMBO_FORMAT_STRING)
    scenes = obs.obs_frontend_get_scene_names()
    
    for scene in scenes:
        obs.obs_property_list_add_string(scene_for_clips, scene, scene)

    obs.source_list_release(scenes)

    obs.obs_properties_add_bool(props, "disabled", "Disable Clipping")
    obs.obs_properties_add_bool(props, "disable_notif", "Disable Notification On Save")

    return props

toaster = WindowsToaster('OBS Replay')
sett = None

def script_load(settings):
    obs.obs_frontend_add_event_callback(obs_frontend_callback)
    obs.timer_add(auto_replay_buffer, 10000)

    global sett
    sett = settings

def script_unload():
    obs.timer_remove(auto_replay_buffer)
    toaster.clear_toasts()

def obs_frontend_callback(event):
    if event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
        path = move_recording()
        if not obs.obs_data_get_bool(sett, "disable_notif"):
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
    if obs.obs_data_get_bool(sett, "disabled"):
        if obs.obs_frontend_replay_buffer_active():
            obs.obs_frontend_replay_buffer_stop()
        return

    scene_as_source = obs.obs_frontend_get_current_scene()

    if obs.obs_source_get_name(scene_as_source) != obs.obs_data_get_string(sett, "scene"):
        obs.obs_source_release(scene_as_source)
        return

    current_scene = obs.obs_scene_from_source(scene_as_source)
    scene_item = obs.obs_scene_find_source_recursive(current_scene, "Fullscreen Game Capture")
    source = obs.obs_sceneitem_get_source(scene_item)

    if (not obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) > 0):
        obs.obs_frontend_replay_buffer_start()
        
    elif (obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) == 0):
        obs.obs_frontend_replay_buffer_stop()

    obs.obs_source_release(scene_as_source)

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