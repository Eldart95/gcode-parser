
# coding: utf-8

# In[3]:


import xlsxwriter


# In[18]:


import sys

workbook = xlsxwriter.Workbook('Parsed_G_Code.xlsx')   
# The workbook object is then used to add new  
# worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet() 



# In[19]:


file = open(sys.argv[1], mode = 'r')
lines = file.readlines()



# In[20]:


e_x = 0
e_y = 1
e_z = 2
row = 1

worksheet.write(0, e_x , 'X')
worksheet.write(0, e_y , 'Y')
worksheet.write(0, e_z , 'Z')

print("Starting to Parse...")
for line in lines:
    if line[0]=='G' and line[1]=='1':
        spl = line.split(" ")
        for string in spl:
            if string[0]=='X':
                worksheet.write(row, e_x , string[1:]) 
            elif string[0]=='Y':
                worksheet.write(row, e_y , string[1:])
            elif string[0]=='Z':
                worksheet.write(row, e_z , string[1:])
    	row+=1            


# In[21]:

print("Done Parsing...")
print("Closing Files..")
workbook.close()


# In[22]:


file.close()
print("Done")
