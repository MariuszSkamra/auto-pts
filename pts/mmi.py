

class MMI:

    MMI_Style_Ok_Cancel1 = 0x11041
    MMI_Style_Ok_Cancel2 = 0x11141
    MMI_Style_Ok = 0x11040
    MMI_Style_Yes_No1 = 0x11044
    MMI_Style_Yes_No_Cancel1 = 0x11043
    MMI_Style_Abort_Retry1 = 0x11042
    MMI_Style_Edit1 = 0x12040
    MMI_Style_Edit2 = 0x12140

    MMI_STYLE_STRING = {
        MMI_Style_Ok_Cancel1: "MMI_Style_Ok_Cancel1",
        MMI_Style_Ok_Cancel2: "MMI_Style_Ok_Cancel2",
        MMI_Style_Ok: "MMI_Style_Ok",
        MMI_Style_Yes_No1: "MMI_Style_Yes_No1",
        MMI_Style_Yes_No_Cancel1: "MMI_Style_Yes_No_Cancel1",
        MMI_Style_Abort_Retry1: "MMI_Style_Abort_Retry1",
        MMI_Style_Edit1: "MMI_Style_Edit1",
        MMI_Style_Edit2: "MMI_Style_Edit2"
    }

    def __init__(self, project_name, wid, test_case, description, style):
        # Remove whitespaces from project and test case name
        project_name = project_name.replace(" ", "")
        test_case = test_case.replace(" ", "")
        pass

