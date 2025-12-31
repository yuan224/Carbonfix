import pywinauto.application
import win32com.client
import time
from pywinauto.keyboard import send_keys
import pywinauto.mouse as mouse

def connect_aspen(file):
    Aspen = win32com.client.Dispatch("Apwn.Document")
    Aspen.InitFromFile2(file)
    Aspen.Visible = 1  # visible
    Aspen.SuppressDialogs = False  # close dialog
    app = pywinauto.application.Application(backend="uia").connect(process=Aspen.ProcessId)
    Asp = app["Dialog"]

    return Aspen, Asp

def reconnect_aspen(Aspen_path):
    print("Aspen crash detected, reconnecting...") 
    try:
        Aspen = win32com.client.Dispatch("Apwn.Document")
        Aspen.InitFromFile2(Aspen_path)    
        Aspen.Visible = 1  # visible
        Aspen.SuppressDialogs = False  # close dialog
        app = pywinauto.application.Application(backend="uia").connect(process=Aspen.ProcessId)
        Asp = app["Dialog"]
        print("Reconnected to Aspen successfully.")
        return Aspen, Asp
    except Exception as e:
        print("Reconnect failed:", e)


def close_aspen(Aspen):
    Aspen.Close()
    
def click_components(Asp):
    Asp.child_window(title="Components", auto_id="igRibbon_btnProperties_Components", control_type="Button").click()
    
def cd_simulation(Asp):
    try:
        Asp.set_focus()
        Asp.children()[3].children()[0].children()[7].select()  # 6=Properties 7=Simulation
    except:
        print("cd_simulation")
        time.sleep(1)
        #pywinauto.mouse.click(coords=(100, 830))

def cd_properties(Asp):
    try:
        Asp.set_focus()
        Asp.children()[3].children()[0].children()[6].select()
    except:
        print("cd_properties")
        time.sleep(1)
        #pywinauto.mouse.click(coords=(100, 780))
        
def click_review(Asp):
    while True:
        try:
            Asp.child_window(
                title="Selection", auto_id="MMTabItem_1", control_type="TabItem"
            ).child_window(auto_id="MMTabPage_1", control_type="Custom").child_window(
                title="cmdReview", auto_id="MMCmdButton_6", control_type="Custom"
            ).child_window(
                title="Review", auto_id="PART_BUTTON", control_type="Button"
            ).click()
        except:
            time.sleep(0.3)
            continue
        break

def reset(Asp):
    Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(auto_id="igRibbon_btnReset", control_type="Button").click()
    # Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(title="Reset",auto_id="igRibbon_btnPrpertiesReset",control_type="Button").click()
    Asp.window(title="Reinitialize").child_window(title="OK",auto_id="Button_1",control_type="Button").click()
    Asp.window(title="Reinitialize").child_window(title="OK",auto_id="btn0",control_type="Button").click()

def run(Asp):
    try:
        Asp.set_focus()
        Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(title="Run", auto_id="igRibbon_btnRunProp", control_type="Button").click()
    except:
        print("run")
        Asp.set_focus()
        #mouse.click(button='left', coords=(190, 8))    
    time.sleep(2)

        
    
def run_sim(Asp):
    try:
        Asp.set_focus()
        Asp.child_window(
        auto_id="igRibbon_QuickAccessToolbar_1",
        control_type="ToolBar"
        ).child_window(
        title="Run",
        auto_id="igRibbon_btnRun",
        control_type="Button"
        ).click()
    except:
        print("run_sim")
        Asp.set_focus()
        #mouse.click(button='left', coords=(230, 75))
    time.sleep(2)
    

def click_Find(asp):
    asp.child_window(title="Find", auto_id="PART_BUTTON", control_type="Button").click()


    
class FindCompounds:
    def __init__(self, asp):
        find_compounds_box = asp.window(title="Find Compounds")
        self.box = find_compounds_box
    
    def input_CAS(self, CAS_number):
        edit_box = self.box.child_window(auto_id="txtContainValue", control_type="Edit")
        edit_box.set_text(CAS_number)
    
    def equal(self):
        equals_radio = self.box.child_window(auto_id="radioEquals", control_type="RadioButton")
        equals_radio.click()
    
    def find_now(self):
        find_now_btn = self.box.child_window(title="Find Now", auto_id="btnFind", control_type="Button")
        find_now_btn.click()
        time.sleep(1.5)
        
    def if_match(self):
        '''
        Check if matching hint is displayed.
        '''
        time.sleep(1.8)
        status_bar = self.box.child_window(auto_id="StatusBar_1", control_type="StatusBar")

        text_ctrl = status_bar.child_window(control_type="Text")
        if text_ctrl.exists():
            sub_status_text = text_ctrl.window_text()
            if 'Matches found: 1' in sub_status_text:
                return True ,'Match'
            
            if 'Matches found' in sub_status_text:
                return True, 'Multiple Matches'
            
            elif 'No Match' in sub_status_text:
                return False, 'No Match'
            
            else:
                return False, 'Not found yet.'
            
        else:
            return False, 'Error'

    def add_comp(self):
        list_item = self.box.child_window(title="AspenTech.AspenPlus.UnmanagedWrapper.DbankComponentItem", control_type="DataItem", found_index=0)
        list_item.select()
        
        add_selected_comp_btn = self.box.child_window(title="Add selected compounds", auto_id="btnAdd", control_type="Button")
        add_selected_comp_btn.click()
    
    def close_fc(self):
        close_btn = self.box.child_window(title="Close", auto_id="btnClose", control_type="Button")
        close_btn.click()
        
def click_comp_cell(Asp,index):
    found_index = index*5-5
    Asp.set_focus()
    cell = Asp.child_window(class_name="GridControlCell", control_type="Custom",found_index=found_index).click_input()
    time.sleep(0.5)

# def click_nrtl():
#     mouse.click(button='left', coords=(115,470))
#     time.sleep(0.2)
#     mouse.click(button='left', coords=(480,105))
#     time.sleep(0.3)
#     mouse.click(button='left', coords=(960,675))
#     time.sleep(2)
#     mouse.click(button='left', coords=(1000,560))
    
def input_CAS_COMP_list(Aspen,Asp, CAS_list):
    results = []
    n=len(CAS_list)
    Asp.set_focus()
    fc = FindCompounds(Asp)
    time.sleep(5)
    Asp.set_focus()
    for i,CAS in enumerate(CAS_list):
        click_comp_cell(Asp, i+1)
        click_Find(Asp)
        fc.equal()
        fc.input_CAS(CAS)
        fc.find_now()
        is_match, status = fc.if_match()
        results.append((i+1, CAS, f"COMP{i+1}", status))
        if is_match:
            fc.add_comp()
        fc.close_fc()
        try:
            Asp.set_focus()
            Asp.child_window(auto_id="chkDontshow",control_type="CheckBox").click_input()
            Asp.child_window(auto_id="btn0",control_type="Button").click_input()
        except:
            pass
    time.sleep(1)    
    for _ in range(10-n):
        Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.RemoveRow(0,n)
    print("\n" + "=" * 65)
    print(f"{'Component':<10} | {'CAS Number':<12} | {'Component ID':<10} | {'Status'}")
    print("=" * 65)
    for comp, CAS, id, status in results:
        print(f"{comp:<10} | {CAS:<12} | {id:<10} | {status}")
    print("=" * 65)
    time.sleep(1)
    Asp.child_window(title="NRTL-1",control_type="TreeItem").click_input()
    for i in range(n):
        Aspen.Application.Tree.FindNode(f"\Data\Properties\Analysis\MIX-1\Input\FLOW\COMP{i+1}").Value = 1
    Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(title="Save",auto_id="igRibbon_ButtonTool_3",control_type="Button").click()
    cd_simulation(Asp)
    Asp.child_window(auto_id="MMTabItem_2", control_type="TabItem").select()
    for i in range(10, n, -1):
            delete_C(Asp, i)
# def delete_C(Asp,n):
#     x = 300
#     base_y = 300
#     step = 25
#     y = base_y + (n - 1) * step
#     Asp.set_focus()
#     time.sleep(0.2)
#     mouse.click(button='left', coords=(x, y))
#     Asp.child_window(title="cmdDelete", control_type="Custom").child_window(auto_id="PART_BUTTON", control_type="Button").click()
#     time.sleep(0.2)
#     Asp.child_window(title="Confirm", control_type="Window").child_window(auto_id="btn0", control_type="Button").click()
#     time.sleep(0.2)
#     Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(title="Save",auto_id="igRibbon_ButtonTool_3",control_type="Button").click()
def delete_C(Asp,n):
    Asp.set_focus()
    time.sleep(0.2)
    cells = Asp.child_window(class_name="GridControlCell",control_type="Custom",found_index=n*2-1)
    cells.click_input()
    Asp.child_window(title="cmdDelete", control_type="Custom").child_window(auto_id="PART_BUTTON", control_type="Button").click()
    time.sleep(0.2)
    Asp.child_window(title="Confirm", control_type="Window").child_window(auto_id="btn0", control_type="Button").click()
    time.sleep(0.2)
    Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", control_type="ToolBar").child_window(title="Save",auto_id="igRibbon_ButtonTool_3",control_type="Button").click()