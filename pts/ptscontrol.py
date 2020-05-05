import ctypes
import pythoncom
import win32com.client
import win32com.server.util
from win32com.server.connect import ConnectableServer
import wmi

PTS_PROC_NAME = "PTS.exe"
PTSCONTROL_PROG_ID = "ProfileTuningSuite_6.PTSControlServer"

PTSCONTROL_E_GUI_UPDATE_FAILED = 0x849C0001
PTSCONTROL_E_PTS_FILE_FAILED_TO_INITIALIZE = 0x849C0002
PTSCONTROL_E_FAILED_TO_CREATE_WORKSPACE = 0x849C0003
PTSCONTROL_E_CLIENT_LOG_NOT_EXPECTED_TO_FAIL = 0x849C0004
PTSCONTROL_E_FAILED_TO_OPEN_WORKSPACE = 0x849C0005
PTSCONTROL_E_PROJECT_NOT_FOUND = 0x849C0010
PTSCONTROL_E_TESTCASE_NOT_FOUND = 0x849C0011
PTSCONTROL_E_TESTCASE_NOT_STARTED = 0x849C0012
PTSCONTROL_E_INVALID_TEST_SUITE = 0x849C0013
PTSCONTROL_E_PTS_VERSION_NOT_FOUND = 0x849C0014
PTSCONTROL_E_PROJECT_VERSION_NOT_FOUND = 0x849C0015
PTSCONTROL_E_TESTCASE_NOT_ACTIVE = 0x849C0016
PTSCONTROL_E_TESTCASE_TIMEOUT = 0x849C0017
PTSCONTROL_E_INVALID_IXIT_PARAM_VALUE = 0x849C0020
PTSCONTROL_E_IXIT_PARAM_NOT_CHANGED = 0x849C0021
PTSCONTROL_E_IXIT_PARAM_UPDATE_FAILED = 0x849C0022
PTSCONTROL_E_IXIT_PARAM_NOT_FOUND = 0x849C0023
PTSCONTROL_E_TEST_SUITE_PARAM_UPDATE_FAILED = 0x849C0024
PTSCONTROL_E_PICS_ENTRY_UPDATE_FAILED = 0x849C0030
PTSCONTROL_E_PICS_ENTRY_NOT_FOUND = 0x849C0031
PTSCONTROL_E_PICS_ENTRY_NOT_CHANGED = 0x849C0032
PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_REGISTERED = 0x849C0040
PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_ALREADY_REGISTERED = 0x849C0041
PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_EXPECTED_TO_FAIL = 0x849C0042
PTSCONTROL_E_BLUETOOTH_ADDRESS_NOT_FOUND = 0x849C0043
PTSCONTROL_E_INTERNAL_ERROR = 0x849C0044
PTSCONTROL_E_FUNCTION_NOT_IMPLEMENTED = 0x849C0099
PTSCONTROL_E_NOINTERFACE = 0x8000400211

PTSCONTROL_E_STRING = {
    PTSCONTROL_E_GUI_UPDATE_FAILED:
        "PTSCONTROL_E_GUI_UPDATE_FAILED",
    PTSCONTROL_E_PTS_FILE_FAILED_TO_INITIALIZE:
        "PTSCONTROL_E_PTS_FILE_FAILED_TO_INITIALIZE",
    PTSCONTROL_E_FAILED_TO_CREATE_WORKSPACE:
        "PTSCONTROL_E_FAILED_TO_CREATE_WORKSPACE",
    PTSCONTROL_E_CLIENT_LOG_NOT_EXPECTED_TO_FAIL:
        "PTSCONTROL_E_CLIENT_LOG_NOT_EXPECTED_TO_FAIL",
    PTSCONTROL_E_FAILED_TO_OPEN_WORKSPACE:
        "PTSCONTROL_E_FAILED_TO_OPEN_WORKSPACE",
    PTSCONTROL_E_PROJECT_NOT_FOUND:
        "PTSCONTROL_E_PROJECT_NOT_FOUND",
    PTSCONTROL_E_TESTCASE_NOT_FOUND:
        "PTSCONTROL_E_TESTCASE_NOT_FOUND",
    PTSCONTROL_E_TESTCASE_NOT_STARTED:
        "PTSCONTROL_E_TESTCASE_NOT_STARTED",
    PTSCONTROL_E_INVALID_TEST_SUITE:
        "PTSCONTROL_E_INVALID_TEST_SUITE",
    PTSCONTROL_E_PTS_VERSION_NOT_FOUND:
        "PTSCONTROL_E_PTS_VERSION_NOT_FOUND",
    PTSCONTROL_E_PROJECT_VERSION_NOT_FOUND:
        "PTSCONTROL_E_PROJECT_VERSION_NOT_FOUND",
    PTSCONTROL_E_TESTCASE_NOT_ACTIVE:
        "PTSCONTROL_E_TESTCASE_NOT_ACTIVE",
    PTSCONTROL_E_TESTCASE_TIMEOUT:
        "PTSCONTROL_E_TESTCASE_TIMEOUT",
    PTSCONTROL_E_INVALID_IXIT_PARAM_VALUE:
        "PTSCONTROL_E_INVALID_IXIT_PARAM_VALUE",
    PTSCONTROL_E_IXIT_PARAM_NOT_CHANGED:
        "PTSCONTROL_E_IXIT_PARAM_NOT_CHANGED",
    PTSCONTROL_E_IXIT_PARAM_UPDATE_FAILED:
        "PTSCONTROL_E_IXIT_PARAM_UPDATE_FAILED",
    PTSCONTROL_E_IXIT_PARAM_NOT_FOUND:
        "PTSCONTROL_E_IXIT_PARAM_NOT_FOUND",
    PTSCONTROL_E_TEST_SUITE_PARAM_UPDATE_FAILED:
        "PTSCONTROL_E_TEST_SUITE_PARAM_UPDATE_FAILED",
    PTSCONTROL_E_PICS_ENTRY_UPDATE_FAILED:
        "PTSCONTROL_E_PICS_ENTRY_UPDATE_FAILED",
    PTSCONTROL_E_PICS_ENTRY_NOT_FOUND:
        "PTSCONTROL_E_PICS_ENTRY_NOT_FOUND",
    PTSCONTROL_E_PICS_ENTRY_NOT_CHANGED:
        "PTSCONTROL_E_PICS_ENTRY_NOT_CHANGED",
    PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_REGISTERED:
        "PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_REGISTERED",
    PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_ALREADY_REGISTERED:
        "PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_ALREADY_REGISTERED",
    PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_EXPECTED_TO_FAIL:
        "PTSCONTROL_E_IMPLICIT_SEND_CALLBACK_NOT_EXPECTED_TO_FAIL",
    PTSCONTROL_E_BLUETOOTH_ADDRESS_NOT_FOUND:
        "PTSCONTROL_E_BLUETOOTH_ADDRESS_NOT_FOUND",
    PTSCONTROL_E_INTERNAL_ERROR:
        "PTSCONTROL_E_INTERNAL_ERROR",
    PTSCONTROL_E_FUNCTION_NOT_IMPLEMENTED:
        "PTSCONTROL_E_FUNCTION_NOT_IMPLEMENTED",
    PTSCONTROL_E_NOINTERFACE:
        "PTSCONTROL_E_NOINTERFACE",
}


class PTSLogger(ConnectableServer):
    """PTS control client logger callback implementation"""

    _reg_desc_ = "AutoPTS Logger"
    _reg_clsid_ = "{50B17199-917A-427F-8567-4842CAD241A1}"
    _reg_progid_ = "autopts.PTSLogger"
    _public_methods_ = ['Log'] + ConnectableServer._public_methods_

    def __init__(self, callback):
        """"Constructor"""

        super(PTSLogger, self).__init__()
        self.callback = callback

    def Log(self, log_type, logtype_string, log_time, log_message):
        """Implements:

        void Log(
                        [in] unsigned int logType,
                        [in] BSTR szLogType,
                        [in] BSTR szTime,
                        [in] BSTR pszMessage);
        };
        """

        self.callback(log_type, logtype_string, log_time, log_message)


class PTSSender(ConnectableServer):
    """PTS control client implicit send callback implementation"""

    _reg_desc_ = "AutoPTS Sender"
    _reg_clsid_ = "{9F4517C9-559D-4655-9032-076A1E9B7654}"
    _reg_progid_ = "autopts.PTSSender"
    _public_methods_ = ['OnImplicitSend'] + ConnectableServer._public_methods_

    def __init__(self, callback):
        """"Constructor"""

        super(PTSSender, self).__init__()
        self.callback = callback

    def OnImplicitSend(self, project_name, wid, test_case, description, style):
        """Implements:

        VARIANT OnImplicitSend(
                        [in] BSTR pszProjectName,
                        [in] unsigned short wID,
                        [in] BSTR szTestCase,
                        [in] BSTR szDescription,
                        [in] unsigned long style);
        };
        """

        rsp = self.callback(project_name, wid, test_case, description, style)
        if rsp:
            is_present = 1
        else:
            is_present = 0

        # Stringify response
        rsp = str(rsp)
        rsp_len = str(len(rsp))
        is_present = str(is_present)

        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
                                       [rsp, rsp_len, is_present])


class PtsControlException(Exception):
    def __init__(self, *args):
        self.ecode = 0
        self.message = ""

        if args:
            if isinstance(args[0], pythoncom.com_error):
                self.ecode = self._com_error_hresult(args[0])
                if self.ecode in PTSCONTROL_E_STRING:
                    self.message = PTSCONTROL_E_STRING[args[0]]
                else:
                    self.message = "COM_ERROR"
            else:
                self.message = args[0]

    def __str__(self):
        return "%d: %s" % (self.ecode, self.message)

    @staticmethod
    def _com_error_hresult(err):
        _, source, description, _, _, hresult = err.excepinfo
        return ctypes.c_uint32(hresult).value


class PtsControl:

    def __init__(self):

        # Get PTS process list before running new PTS daemon
        c = wmi.WMI()
        pts_ps_list_pre = []
        pts_ps_list_post = []

        for ps in c.Win32_Process(name=PTS_PROC_NAME):
            pts_ps_list_pre.append(ps)

        self._pts = win32com.client.Dispatch(PTSCONTROL_PROG_ID)

        # Get PTS process list after running new PTS daemon to get PID of
        # new instance
        for ps in c.Win32_Process(name=PTS_PROC_NAME):
            pts_ps_list_post.append(ps)

        pts_ps_list = list(set(pts_ps_list_post) - set(pts_ps_list_pre))
        if not pts_ps_list:
            return

        self._pts_proc = pts_ps_list[0]

        self._on_log_callback = None
        self._on_implicit_send_callback = None

        self._pts_sender = win32com.client.dynamic.Dispatch(
            win32com.server.util.wrap(PTSSender(self._on_implicit_send)))
        try:
            self._pts.RegisterImplicitSendCallbackEx(self._pts_sender)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def create_workspace(self):
        pass

    def open_workspace(self, workspace_path):
        try:
            return self._pts.OpenWorkspace(workspace_path)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def get_project_count(self):
        try:
            return self._pts.GetProjectCount()
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def get_project_name(self, project_index):
        try:
            return self._pts.GetProjectName(project_index)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def get_project_version(self, project_name):
        try:
            return self._pts.GetProjectVersion(project_name)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def get_test_case_count(self, project_name):
        try:
            return self._pts.GetTestCaseCount(project_name)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def get_test_case_description(self, project_name, test_case_index):
        try:
            return self._pts.GetTestCaseDescription(project_name,
                                                    test_case_index)
        except pythoncom.com_error as err:
            raise PtsControlException(err)

    def is_active_test_case(self, project_name, test_case_name):
        return self._pts.IsActiveTestCase(project_name, test_case_name)

    def run_test_case(self, project_name, test_case_name):
        return self._pts.RunTestCase(project_name, test_case_name)

    def stop_test_case(self):
        return self._pts.StopTestCase()

    def get_test_case_count_from_tss_file(self, project_name):
        return self._pts.GetTestCaseCountFromTSSFile(project_name)

    def get_test_cases_from_tss_file(self, project_name):
        return self._pts.GetTestCasesFromTSSFile(project_name)

    def update_pics(self, project_name, entry_name, bool_value):
        return self._pts.UpdatePics(project_name, entry_name, bool_value)

    def update_pixit_param(self, project_name, param_name, param_value):
        return self._pts.UpdatePixitParam(project_name, param_name, param_value)

    def set_control_client_logger_callback(self, callback):
        return self._pts.SetControlClientLoggerCallback(callback)

    def register_implicit_send_callback_ex(self, callback):
        return self._pts.RegisterImplicitSendCallbackEx(callback)

    def unregister_implicit_send_callback_ex(self):
        return self._pts.UnregisterImplicitSendCallbackEx()

    def enable_maximum_logging(self, enable):
        return self._pts.EnableMaximumLogging(enable)

    def set_pts_call_timeout(self, timeout):
        return self._pts.SetPTSCallTimeout(timeout)

    def save_test_history_log(self, save):
        return self._pts.SaveTestHistoryLog(save)

    def get_pts_bluetooth_address(self):
        return self._pts.GetPTSBluetoothAddress()

    def get_pts_version(self):
        return self._pts.GetPTSVersion()
