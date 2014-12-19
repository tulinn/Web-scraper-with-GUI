import urllib2
import xlwt
import Tkinter
from Tkinter import *
import tkMessageBox

from bs4 import BeautifulSoup

top = Tkinter.Tk()
#Code to add widgets will go here

Label2 = Label(top, text="Street Name, City, Number of Entries (ex. Yonge St, Toronto, 15):")
#E1 = Entry(top, bd = 5)

#E1.pack(side = RIGHT)
global streets_list
streets_list = []
v = StringVar()
entry1 = Entry(top, textvariable = v)

def Added():
    """
    Collecting the user set variabled for the Running function to execute
    """
    global num
    st = v.get() 
    listof = st.split()
    num = int(listof.pop())
    string_1 = st.find(',')
    string_2 = st.find(',', string_1+1)
    
    st = st[0:string_2]
        
    st = st.title()
    street_new = st
    if st.find(',') != -1:
        st = st.replace(" ", "+")
        streets_list.append(st) 
        v.set("")    
        tkMessageBox.showinfo("Add The Street Names", street_new + " " + "Added!")
    else:
        tkMessageBox.showinfo("Add The Street Names", "Please enter the correct format!")
        
def Running():
    """
    Makes use of the previously collected user information from Added function
    """
    #tkMessageBox.showinfo("Add The Street Names", "Running...")   
    
    def setup():
        global temp_link, soup
        temp_link = "http://411.ca/search/?lang=en&q={0}&st=reverse&p={1}&".format(street, page) 
        request = urllib2.Request(temp_link)
        response = urllib2.urlopen(request)
        soup = BeautifulSoup(response)      
    
    for street in streets_list:
        workbook = xlwt.Workbook(encoding="utf-8")
        sheet1 = workbook.add_sheet("Sheet 1")
        current_row_count = 0
        sheet1.write(current_row_count, 0, 'Street')
        sheet1.write(current_row_count, 1, 'Address')
        sheet1.write(current_row_count, 2, 'City')
        sheet1.write(current_row_count, 3, 'PC')
        sheet1.write(current_row_count, 4, 'Name')
        sheet1.write(current_row_count, 5, 'Phone #')
        current_row_count += 1     
        profile_street = street.replace("+", " ")
        sheet1.write(current_row_count, 0, profile_street)    
        street = street.replace(" ", "+")
        #print "current street: " + profile_street 
        try:
            page = 1
            setup()
            profile_count = 0
            street_flag = True  #to make sure we add the street name only once
            while (soup.title.string != '411.ca Reverse Lookup'):
                #print page
                #operation - begin
                page_divs = soup.findAll("div", { "class" : "founditem_content reverse" })
                for profile in page_divs:
                    # First break statement, second one should be identical
                    if profile_count > num: 
                        break
                    profile_llogo = profile.find("div", {"class" : "llogo"})
                    profile_link = profile_llogo.contents[1]['content']
                    if ('person' in profile_link):
                        #continue operation - it is a person profile
                        profile_ltitle = profile.find("div", {"class" : "ltitle"})
                        profile_name = profile_ltitle.contents[1].string
                        profile_name = profile_name.replace("   ", " ")
                        profile_name = profile_name.replace("  ", " ")
                        profile_linfo = profile.find("div", {"class" : "linfo"})
                        profile_phone = profile_linfo.contents[7].string
                        profile_full_address = profile_linfo.contents[3].string.replace("\n", "")
                        profile_full_address = profile_full_address.replace("\t", "") 
                        profile_address = ""
                        profile_city = ""
                        profile_pc = "" 
                        comma_count = 0
                        for character in profile_full_address:
                            if (character == ','):
                                comma_count += 1
                            if (character != ","):
                                if (comma_count == 0):  #for the profile_address
                                    if (character == " " and profile_address):
                                        profile_address += character 
                                    if (character != " "):
                                        profile_address += character
                                elif (comma_count == 1): #for the profile_city
                                    if (character == " " and profile_city):
                                        profile_city += character 
                                    if (character != " "):
                                        profile_city += character
                                elif (comma_count == 3): #for profile_pc
                                    if (character == " " and profile_pc):
                                        profile_pc += character 
                                    if (character != " "):
                                        profile_pc += character
                        #at this point we have -> street, profile_name, profile_phone, profile_address, profile_city, profile_pc
                        #if (street_flag): #add the street name to the database only once
                        try:
                            #cur.execute('INSERT INTO "411ca-2" (address, city, pc, name, phone) VALUES (%s, %s, %s, %s, %s);',
                                        #[profile_address, profile_city, profile_pc, profile_name, profile_phone]) 
                            #write address, city, pc, name, phone
                            sheet1.write(current_row_count, 1, profile_address)
                            sheet1.write(current_row_count, 2, profile_city)
                            sheet1.write(current_row_count, 3, profile_pc)
                            sheet1.write(current_row_count, 4, profile_name)
                            sheet1.write(current_row_count, 5, profile_phone)                
                            current_row_count += 1
                            profile_count += 1
                        except Exception,e:
                            print e
                        #end the operation for this current profile
                    else:
                        continue #go on to the next profile in page_divs
                    #operation - end
                # Second break statement, first one should be identical
                if profile_count > num:
                    break
                page += 1
                setup()
        except Exception,e:
            print e
            print temp_link
            continue 
        
        workbook.save(profile_street + ".xls")
        
    #print 'Operation done'
    tkMessageBox.showinfo("Add The Street Names", "Done!")
      
    
Button1 = Tkinter.Button(top, text = "Run", command = Running)
Button1.grid(row=0, column=0)

Button2 = Tkinter.Button(top, text = "Add", command = Added)
Button2.grid(row=0, column=10)

Button1.pack(side = BOTTOM)
Button2.pack(side = BOTTOM)

Label2.pack(side = LEFT) 
entry1.pack(side = RIGHT)

top.mainloop()

