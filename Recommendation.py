# import openpyxl and tkinter modules
import json

from openpyxl import *
from tkinter import *

# globally declare wb and sheet variable

# opening the existing excel file
wb = load_workbook('./excel.xlsx')

# create the sheet object
sheet = wb.active


# Function to set focus (cursor)
def focus1(event):
    Class_field.focus_set()
# Function to set focus
def focus2(event):
    # set focus on the sem_field box
    OperationCode_field.focus_set()
# Function to set focus
def focus3(event):
    # set focus on the form_no_field box
    Description_field.focus_set()
# Function to set focus
def focus4(event):
    # set focus on the contact_no_field box
    Image_field.focus_set()
# Function to set focus
def focus5(event):
    # set focus on the email_id_field box
    FlatLaborHour_field.focus_set()
def focus6(event):
    # set focus on the email_id_field box
    FlatLaborHour_field.focus_set()
# Function to set focus
def focus7(event):
    # set focus on the address_field box
    StartDate_field.focus_set()
def focus8(event):
    # set focus on the address_field box
    EndDate_field.focus_set()
def focus9(event):
    # set focus on the address_field box
    Make_field.focus_set()
def focus10(event):
    # set focus on the address_field box
    Active_field.focus_set()
def focus11(event):
    # set focus on the address_field box
    DealerID_field.focus_set()
def focus12(event):
    # set focus on the address_field box
    LastUpdatedByUser_field.focus_set()
def focus13(event):
    # set focus on the address_field box
    LastUpdatedDateTime_field.focus_set()
def focus14(event):
    # set focus on the address_field box
    LastUpdatedByDisplayName_field.focus_set()
def focus15(event):
    # set focus on the address_field box
    CreateDateTime_field.focus_set()
def focus16(event):
    # set focus on the address_field box
    DocumentVersion_field.focus_set()



# Function for clearing the
# contents of text entry boxes
def clear():
    # clear the content of text entry box
    Class_field.delete(0, END)
    OperationCode_field.delete(0, END)
    Description_field.delete(0, END)
    Image_field.delete(0, END)
    FlatLaborHour_field.delete(0, END)
    FlatLaborHour_field.delete(0, END)
    StartDate_field.delete(0, END)
    EndDate_field.delete(0, END)
    Make_field.delete(0, END)
    Active_field.delete(0, END)
    DealerID_field.delete(0, END)
    LastUpdatedByUser_field.delete(0, END)
    LastUpdatedDateTime_field.delete(0, END)
    LastUpdatedByDisplayName_field.delete(0, END)
    CreateDateTime_field.delete(0, END)
    DocumentVersion_field.delete(0, END)

def createJSON():

    file=dict({'class':Class_field.get(),'operationCode':OperationCode_field.get(),'description':Description_field.get(),
               'image':Image_field.get(),'flatLaborHour':FlatLaborHour_field.get(),'flatLaborPrice':FlatLaborPrice_field.get(),
               'startDate':StartDate_field.get(),'endDate':EndDate_field.get(),'make':Make_field.get(),'active':Active_field.get(),
               'dealerID':DealerID_field.get(),'lastUpdatedByUser':LastUpdatedByUser_field.get(),'lastUpdatedDateTime':LastUpdatedDateTime_field.get(),
               'lastUpdatedByDisplayName':LastUpdatedByDisplayName_field.get(),'createDateTime':CreateDateTime_field.get(),'documentVersion':DocumentVersion_field.get()})
    y = json.dumps(file)
    print(y)


if __name__ == "__main__":
    # create a GUI window
    root = Tk()

    # set the background colour of GUI window
    root.configure(background='grey')

    # set the title of GUI window
    root.title("Recommendations")

    # set the configuration of GUI window
    root.geometry("1000x800")

   # excel()

    # create a Form label
    heading = Label(root, text="Dealer Recommendation", bg="grey")

    result = Label(root, text="Result", bg="grey").grid(row=0,column=3)

    Class = Label(root, text="Class", bg="grey")

    OperationCode = Label(root, text="OperationCode", bg="grey")

    Description = Label(root, text="Description", bg="grey")

    Image = Label(root, text="Image", bg="grey")

    FlatLaborHour = Label(root, text="FlatLaborHour", bg="grey")

    FlatLaborPrice = Label(root, text="FlatLaborPrice", bg="grey")

    StartDate = Label(root, text="StartDate", bg="grey")

    EndDate = Label(root, text="EndDate", bg="grey")

    Make = Label(root, text="ApplicableForYearMakeModel", bg="grey")

    Active = Label(root, text="IsActive", bg="grey")

    DealerID = Label(root, text="DealerID", bg="grey")

    LastUpdatedByUser = Label(root, text="LastUpdatedByUser", bg="grey")

    LastUpdatedDateTime = Label(root, text="LastUpdatedDateTime", bg="grey")

    LastUpdatedByDisplayName = Label(root, text="LastUpdatedByDisplayName", bg="grey")

    CreateDateTime = Label(root, text="CreateDateTime", bg="grey")

    DocumentVersion = Label(root, text="DocumentVersion", bg="grey")


    # grid method is used for placing
    # the widgets at respective positions
    # in table like structure .
    heading.grid(row=0, column=1)
    Class.grid(row=1, column=0)
    OperationCode.grid(row=2, column=0)
    Description.grid(row=3, column=0)
    Image.grid(row=4, column=0)
    FlatLaborHour.grid(row=5, column=0)
    FlatLaborPrice.grid(row=6, column=0)
    StartDate.grid(row=7, column=0)
    EndDate.grid(row=8, column=0)
    Make.grid(row=9, column=0)
    Active.grid(row=10, column=0)
    DealerID.grid(row=11, column=0)
    LastUpdatedByUser.grid(row=11, column=0)
    LastUpdatedDateTime.grid(row=12, column=0)
    LastUpdatedByDisplayName.grid(row=13, column=0)
    CreateDateTime.grid(row=14, column=0)
    DocumentVersion.grid(row=15, column=0)


    # create a text entry box
    # for typing the information
    result=Entry(root)
    Class_field = Entry(root)
    OperationCode_field = Entry(root)
    Description_field = Entry(root)
    Image_field = Entry(root)
    FlatLaborHour_field = Entry(root)
    FlatLaborPrice_field = Entry(root)
    StartDate_field = Entry(root)
    EndDate_field = Entry(root)
    Make_field = Entry(root)
    Active_field = Entry(root)
    DealerID_field = Entry(root)
    LastUpdatedByUser_field = Entry(root)
    LastUpdatedDateTime_field = Entry(root)
    LastUpdatedByDisplayName_field = Entry(root)
    CreateDateTime_field = Entry(root)
    DocumentVersion_field = Entry(root)



    Class_field.bind("<Return>", focus1)
    OperationCode_field.bind("<Return>", focus2)
    Description_field.bind("<Return>", focus3)
    Image_field.bind("<Return>", focus4)
    FlatLaborHour_field.bind("<Return>", focus5)
    FlatLaborPrice_field.bind("<Return>", focus6)
    StartDate_field.bind("<Return>", focus6)
    EndDate_field.bind("<Return>", focus7)
    Make_field.bind("<Return>", focus8)
    Active_field.bind("<Return>", focus9)
    DealerID_field.bind("<Return>", focus10)
    LastUpdatedByUser_field.bind("<Return>", focus11)
    LastUpdatedDateTime_field.bind("<Return>", focus12)
    LastUpdatedByDisplayName_field.bind("<Return>", focus13)
    CreateDateTime_field.bind("<Return>", focus14)
    DocumentVersion_field.bind("<Return>", focus15)




    # grid method is used for placing
    # the widgets at respective positions
    # in table like structure .
    result.grid(row=1, column=3,ipadx="100",)
    Class_field.grid(row=1, column=1, ipadx="100")
    OperationCode_field.grid(row=2, column=1, ipadx="100")
    Description_field.grid(row=3, column=1, ipadx="100")
    Image_field.grid(row=4, column=1, ipadx="100")
    FlatLaborHour_field.grid(row=5, column=1, ipadx="100")
    FlatLaborPrice_field.grid(row=6, column=1, ipadx="100")
    StartDate_field.grid(row=7, column=1, ipadx="100")
    EndDate_field.grid(row=8, column=1, ipadx="100")
    Make_field.grid(row=9, column=1, ipadx="100")
    Active_field.grid(row=10, column=1, ipadx="100")
    DealerID_field.grid(row=11, column=1, ipadx="100")
    LastUpdatedByUser_field.grid(row=12, column=1, ipadx="100")
    LastUpdatedDateTime_field.grid(row=13, column=1, ipadx="100")
    LastUpdatedByDisplayName_field.grid(row=14, column=1, ipadx="100")
    CreateDateTime_field.grid(row=15, column=1, ipadx="100")
    DocumentVersion_field.grid(row=16, column=1, ipadx="100")



    # call excel function

    # create a Submit Button and place into the root window
    #submit = Button(root, text="Submit", fg="Black",
    #                bg="Red", command=createJSON)
    #submit.grid(row=17, column=1)
    i=1
    while i==1:
        submit = Button(root, text="Submit", fg="Black",
                        bg="Red", command=createJSON)
        submit.grid(row=17, column=1)
        root.mainloop()

    # start the GUI

