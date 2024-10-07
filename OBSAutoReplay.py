import obspython as obs # type: ignore
from datetime import datetime
import re
import os
import psutil
import traceback
import time
import win32api
import win32gui
import win32process
from windows_toasts import Toast, WindowsToaster, ToastDuration

def script_description():
    return """
Quality of Life features for Replay Buffer, making it similar to applications like Nvidia Shadowplay.

OBSAutoReplay v1.1 by Dyvinia
""".strip()

def script_properties():
    props = obs.obs_properties_create()
    scene_for_clips = obs.obs_properties_add_list(props, "scene", "Scene For Clipping",
                                    obs.OBS_COMBO_TYPE_LIST,
                                    obs.OBS_COMBO_FORMAT_STRING)
    scenes = obs.obs_frontend_get_scene_names()
    
    for scene in scenes:
        obs.obs_property_list_add_string(scene_for_clips, scene, scene)

    obs.source_list_release(scenes)
    
    refresh_interval = obs.obs_properties_add_float(props, "refresh_interval", "Refresh Interval:", 1, 20, 1)
    obs.obs_property_set_long_description(refresh_interval, "How often the OBSAutoReplay checks for if a game has started\n**Changing this requires reloading scripts or restarting OBS**")
    
    toast_duration = obs.obs_properties_add_float_slider(props, "toast_duration", "Notification Duration:", 0.5, 5, 0.05)
    obs.obs_property_set_long_description(toast_duration, "How long the notifications stay on screen")
    
    enabled = obs.obs_properties_add_bool(props, "enabled", "Enable Clipping ")
    #obs.obs_property_set_long_description(enabled, "Enable/Disable Clipping")
    
    enable_notif = obs.obs_properties_add_bool(props, "enable_notif", "Notification On Save ")
    obs.obs_property_set_long_description(enable_notif, "Shows a windows notification whenever a clip is saved")

    return props

toaster = WindowsToaster('OBSAutoReplay')
sett = None
hotkey_id = obs.OBS_INVALID_HOTKEY_ID
current_game = None
start_time = None

def script_load(settings):
    obs.obs_frontend_add_event_callback(obs_frontend_callback)
    
    global sett
    sett = settings
    
    obs.timer_add(auto_replay_buffer, int(obs.obs_data_get_double(sett, "refresh_interval") * 1000))

    global hotkey_id
    hotkey_id = obs.obs_hotkey_register_frontend("query_clipping", "Check If Replay Buffer Is Enabled", query_clipping_hotkey)
    the_data_array = obs.obs_data_get_array(settings, "query_clipping")
    obs.obs_hotkey_load(hotkey_id, the_data_array)
    obs.obs_data_array_release(the_data_array)

def script_save(settings):
    the_data_array = obs.obs_hotkey_save(hotkey_id)
    obs.obs_data_set_array(settings, "query_clipping", the_data_array)
    obs.obs_data_array_release(the_data_array)
    
def script_defaults(settings):
    obs.obs_data_set_default_double(settings, "refresh_interval", 10)
    obs.obs_data_set_default_double(settings, "toast_duration", 1.5)
    obs.obs_data_set_default_bool(settings, "enabled", True)
    obs.obs_data_set_default_bool(settings, "enable_notif", True)

def script_unload():
    obs.timer_remove(auto_replay_buffer)
    obs.obs_hotkey_unregister("query_clipping")
    toaster.clear_toasts()

def obs_frontend_callback(event):
    if event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
        # sometimes the moving takes a while and theres no notif until after its done, leaving me worried it didnt save the clip. hopefully this fixes that
        if obs.obs_data_get_bool(sett, "enable_notif"):
            newToast = Toast()
            newToast.text_fields = ['Saving Replay...']
            newToast.duration = ToastDuration.Short
            toaster.clear_toasts()
            toaster.show_toast(newToast)
            time.sleep(obs.obs_data_get_double(sett, "toast_duration"))
            toaster.clear_toasts()
        path = move_recording()
        if obs.obs_data_get_bool(sett, "enable_notif"):
            newToast = Toast()
            newToast.text_fields = ['Saved Replay', "Saved in " + path]
            newToast.duration = ToastDuration.Short
            toaster.clear_toasts()
            toaster.show_toast(newToast)
            time.sleep(obs.obs_data_get_double(sett, "toast_duration"))
            toaster.clear_toasts()
            
    elif event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_STARTED:
        newToast = Toast()

        global current_game
        if current_game:
            newToast.text_fields = ['Started Replay Buffer', "Playing " + current_game]
        else:
            newToast.text_fields = ['Started Replay Buffer']
        
        newToast.duration = ToastDuration.Short
        toaster.clear_toasts()
        toaster.show_toast(newToast)
        
        time.sleep(obs.obs_data_get_double(sett, "toast_duration") * 1.5)
        toaster.clear_toasts()
        
    elif event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_STOPPED:
        newToast = Toast()
        
        global start_time
        if start_time:
            newToast.text_fields = ['Stopped Replay Buffer', "Session Duration: " + get_session_duration()]
        else:
            newToast.text_fields = ['Stopped Replay Buffer']
        
        newToast.duration = ToastDuration.Short
        toaster.clear_toasts()
        toaster.show_toast(newToast)
        
        start_time = None
        # no toast duration here so that the notif with the playtime isnt lost

def auto_replay_buffer():
    if not obs.obs_data_get_bool(sett, "enabled"):
        if obs.obs_frontend_replay_buffer_active():
            obs.obs_frontend_replay_buffer_stop()
        return

    try:
        scene_as_source = obs.obs_frontend_get_current_scene()

        if obs.obs_source_get_name(scene_as_source) != obs.obs_data_get_string(sett, "scene"):
            obs.obs_source_release(scene_as_source)
            return

        scene_items = obs.obs_scene_enum_items(obs.obs_scene_from_source(scene_as_source))

        source = None
        for item in scene_items:
            source_item = obs.obs_sceneitem_get_source(item)
            source_id = obs.obs_source_get_id(source_item)
            if source_id == "game_capture":
                source = source_item
        obs.sceneitem_list_release(scene_items)
        
        if source is None:
            print("Could not find Game Capture source in current scene")

        global current_game
        global start_time
        if (not obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) > 0):
            obs.obs_frontend_replay_buffer_start()
            current_game = get_foreground_window()
            start_time = datetime.now()

        elif (obs.obs_frontend_replay_buffer_active() and obs.obs_source_get_width(source) == 0):
            obs.obs_frontend_replay_buffer_stop()
            current_game = None
            start_time = None

        obs.obs_source_release(scene_as_source)
        
    except Exception:
        traceback.print_exc()

def move_recording():
    replay_buffer = obs.obs_frontend_get_replay_buffer_output()
    cd = obs.calldata_create()
    ph = obs.obs_output_get_proc_handler(replay_buffer)
    obs.proc_handler_call(ph, "get_last_replay", cd)
    path = obs.calldata_string(cd, "path")
    obs.calldata_destroy(cd)
    obs.obs_output_release(replay_buffer)

    global current_game
    if current_game is None:
        current_game = get_foreground_window()
    new_path = os.path.dirname(path) + '/Replays/' + current_game + '/' + os.path.basename(path)

    print("Saving replay to: " + new_path)
    os.renames(path, new_path)

    return '/Replays/' + current_game + '/'

def get_foreground_window():
    try: 
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)

        exe = psutil.Process(pid).exe()
        print(exe)

        try:
            language, codepage = win32api.GetFileVersionInfo(exe, '\\VarFileInfo\\Translation')[0]
            stringFileInfo = u'\\StringFileInfo\\%04X%04X\\%s' % (language, codepage, "FileDescription")

            return alphanumeric(win32api.GetFileVersionInfo(exe, stringFileInfo))
        except:
            return psutil.Process(pid).name().replace(".exe", "").replace('.', '').strip()

    except:
        return "Other"
    
def alphanumeric(s):
    fixed = re.sub(r'[^A-Za-z0-9 ]+', '', s).strip()
    if len(fixed) > 0:
        return fixed
    else:
        return "Other"
    
def query_clipping_hotkey(is_pressed):
    toasterQuery = WindowsToaster('OBSAutoReplay')
    if is_pressed and obs.obs_frontend_replay_buffer_active():
        newToast = Toast()
        
        global current_game
        global start_time
        if current_game and start_time:
            newToast.text_fields = ['Replay Buffer is Currently Active',  f"Playing {current_game} | Session Duration: {get_session_duration()}"]
        elif current_game:
            newToast.text_fields = ['Replay Buffer is Currently Active',  f"Playing {current_game}"]
        elif start_time:
            newToast.text_fields = ['Replay Buffer is Currently Active',  f"Session Duration: {get_session_duration()}"]
        else:
            newToast.text_fields = ['Replay Buffer is Currently Active']
            
        newToast.duration = ToastDuration.Short
        toasterQuery.clear_toasts()
        toasterQuery.show_toast(newToast)
        
        time.sleep(obs.obs_data_get_double(sett, "toast_duration"))
        toasterQuery.clear_toasts()
        
    elif is_pressed and not obs.obs_frontend_replay_buffer_active():
        newToast = Toast()
        newToast.text_fields = ['Replay Buffer is Not Active']
        newToast.duration = ToastDuration.Short
        toasterQuery.clear_toasts()
        toasterQuery.show_toast(newToast)
        
        time.sleep(obs.obs_data_get_double(sett, "toast_duration"))
        toasterQuery.clear_toasts()
        
def get_session_duration():
    global start_time
    return str(datetime.now() - start_time).split('.')[0]