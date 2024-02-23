import win32com.client
from tkinter import *
from tkinter import ttk, filedialog
import os, glob
import pandas as pd
import calendar

# establish connection
ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0

# GUI functions
def open_file():
    folder = filedialog.askdirectory()
    if folder:
        folder_name = folder.title()
        # path format may need to be changed? from \ to /?
        folder_path = os.path.abspath(folder_name)
        # get yr and m info, and change dir to current month's folder
        yr_txt = yr_box.get()
        m_txt = m_box.get()
        # goes into the monthly folder automatically after selecting site folder
        os.chdir(folder_path)
        f_path2 = str(os.getcwd())
        os.chdir('{0}.{1}'.format(yr_txt, m_txt))
        current_path = os.getcwd()
        print(os.getcwd())
        site_name = folder_name.split('/')[-1].upper()
        # folder name for future use
        site_folder = site_name
        # print(site_folder)

        if '-' in site_name and site_name.split('-')[0][0] in '1234567890':
            site_name = site_name.split('-')[1:]
            site_name = '-'.join(site_name)
        l_folder.config(text=site_name)
        print(site_name)
        [i_to, i_cc, i_excel, i_other, i_title] = find_in_lst(site_name)

        # add attachment invoice of this site to email
        print(os.getcwd())
        exist_pdf = str(glob.glob('*.pdf')[0])
        pdf_lst = [exist_pdf]
        print(exist_pdf)

        # add excel scenario
        if i_excel != 0:
            exist_excel = str(glob.glob('*.xlsx')[0])
            newmail.Attachments.Add(current_path + '/' + exist_excel)

        # multiple sites scenario
        if i_other != 0:
            other_lst = i_other
            print(other_lst)
            for i in range(len(other_lst)):
                # TBD HERE: fix path!
                partone = '/'.join(current_path.split('\\')[:-2]) + '/'
                parttwo = '/{0}.{1}'.format(yr_txt, m_txt)
                jjj = partone + other_lst[i] + parttwo

                os.chdir(jjj)
                current_path2 = os.getcwd()
                print(os.getcwd())
                exist_pdf_t = str(glob.glob('*.pdf')[0])
                pt = current_path2 + '/' + exist_pdf_t
                newmail.Attachments.Add(pt)
                pdf_lst.append(exist_pdf_t)
        numba = ''
        for i in pdf_lst:
            numba += '#{} '.format(i.split('_')[0])
        # use ; for multiple recipients!
        newmail.To = i_to
        if i_cc != 0:
            newmail.CC = i_cc+';sample@gmail.com'
        else:
            newmail.CC = 'sample@gmail.com'
        if i_title != 0:
            newmail.Subject = 'Security Service Invoice of {0} {1} {2} {3}'.format(calendar.month_abbr[int(m_txt)],
                                                                                   yr_txt,
                                                                                   numba,
                                                                                   i_title)
        else:
            newmail.Subject = 'Security Service Invoice of {0} {1} {2} {3}'.format(calendar.month_abbr[int(m_txt)],
                                                                                   yr_txt,
                                                                                   numba,
                                                                                   site_name)
        p = current_path + '/' + exist_pdf
        newmail.Attachments.Add(p)
        out = read_from_txt("***Your Path To email_body.txt***")
        newmail.Body = out.format(calendar.month_abbr[int(m_txt)])
    return

def e_preview():
    newmail.Display()
    return


def e_send():
    newmail.Send()
    return


def refresh():
    return


def e_format():
    return


def find_in_lst(site):
    # input: site_name
    # output: [to], [cc], excel_file (if any), [sites send together] (if any), special title (if any)
    # import list of site contacts/info
    f = "***Path to email_lst.xlsx***"
    # create a pandas dataframe
    df = file_import(f)
    # filter out all info based on site name (EXACT MATCH!)
    # the last three options would have 0 as values if no data is given
    i_row = df[df['site'] == site]
    i_to = list(i_row['to'])[0]
    i_cc = list(i_row['cc'].fillna(0))[0]
    i_excel = list(i_row['if_excel'].fillna(0))[0]
    i_other = list(i_row['other_sites'].fillna(0))[0]
    i_title = list(i_row['special_title'].fillna(0))[0]
    if i_other == 0:
        i_other = []
    else:
        i_other = i_other.split(';')
    return [i_to, i_cc, i_excel, i_other, i_title]


def new_email():
    # creating customized email based on invoice name and client info
    # client info: get from excel! email_lst.xlsx
    global newmail
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
    l_folder.config(text='New email created')
    return


def file_import(f_path):
    f_path = f_path
    pwd = os.getcwd()
    os.chdir(os.path.dirname(f_path))
    f = pd.read_excel(os.path.basename(f_path))
    os.chdir(pwd)
    return f
    # print(f)
    print('-----')
    print('File Imported! ')
    # print(f.columns.values)
    print('-----')


def read_from_txt(file):
    f = open(file, "r", encoding='utf-8')
    data = f.read()
    # out = data.split("\n")
    # print('Site imported: ')
    # print(out)
    f.close()
    return data


# path constants


# GUI Interface
root = Tk()
root.geometry("600x400")
root.title('Invoice Email Auto - By Klaus')

# entering info through input box
yr_box = Entry(root, width=20, font=("Comic Sans MS", 12))
m_box = Entry(root, width=20, font=("Comic Sans MS", 12))
yr_box.grid(row=1, column=1)
m_box.grid(row=1, column=2)

# label for showing current path
new_email_button = Button(root, text='New Email', font=("Comic Sans MS", 14),
                          command=new_email)
new_email_button.grid(row=2, column=1)
# loc_input = Label(root, text='Select Path', font=("Comic Sans MS", 14))
# loc_input.grid(row=2, column=1)

# label for showing folder path
l_folder = Label(root, text='Multiple Sites?', font=("Comic Sans MS", 14))
l_folder.grid(row=3, column=1)

# button for refreshing options
path_button = Button(root, text='Select',
                    command=open_file,
                    font=("Comic Sans MS", 12))
path_button.grid(row=2, column=2)

drop_box = ttk.Combobox(state="readonly",
                        values=[])
drop_box.grid(row=3, column=2)

# button for previewing and sending email
preview_button = Button(root, text='Preview', command=e_preview, font=("Comic Sans MS", 12))
preview_button.grid(row=4, column=1)

send_button = Button(root, text='Send', command=e_send, font=("Comic Sans MS", 12))
send_button.grid(row=4, column=2)

# title_text = Label(root, text='Enter the 2 time in box below: ', font=("Comic Sans MS", 16))
# title_text.pack()
#
# enter_box = Entry(root, width=20, font=("Comic Sans MS", 12))
# enter_box.pack()
#
# button = Button(root, text='Convert', command=cal_t_b, pady=10,
# 	padx=20, font=("Comic Sans MS", 14, 'bold'))
# button.pack()
#
# info_l = Label(root, text='Result will be shown here! ', font=("Comic Sans MS", 12))
# info_l.pack()
#
# root.bind('<Return>', cal_t)

root.mainloop()
