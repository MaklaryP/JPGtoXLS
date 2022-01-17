from PIL import Image
import numpy as np
import xlsxwriter as xl 

#XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
#v tejeto casti je mozne menit udaje
nazov_obr = 'test_pic_4.JPG'
nazov_tbl = 'test_pic_4.xlsx'
max_pix = 400
#XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



img = Image.open(nazov_obr)

#zmena obrazka na potrebnu velkost
width, height = img.size

if width > height:
    nw = max_pix
    nh = int(height *(max_pix/width))
else:
    nh = max_pix
    nw = int(width *(max_pix/height))

#otvorenie obrazka
new_img = img.resize( (nw,nh) ) 
new_img.save('new_{}'.format(nazov_obr))

img = new_img

#obrazok do array
ary = np.array(img) #(80, 45, 3)
r,g,b = np.split(ary,3, axis=2)

shape = ary.shape
poc_riad = shape[0]
poc_buniek = shape[1]

workbook = xl.Workbook(nazov_tbl)
worksheet = workbook.add_worksheet()

#cervene policka
for j in range(poc_riad):
    riadok = [int(bunka) for bunka in r[j]]
    for i in range(len(riadok)):
        worksheet.write(0 + 3*j, i, riadok[i])

    worksheet.conditional_format(0 + 3*j,0,0 + 3*j, poc_buniek, {'type':'2_color_scale',
                                                                 'min_color':'#000000',
                                                                 'max_color':'#ff0000',
                                                                 'min_value':0,
                                                                 'max_value':255})

#zelene policka

for j in range(poc_riad):
    riadok = [int(bunka) for bunka in g[j]]
    for i in range(len(riadok)):
        worksheet.write(1 + 3*j, i, riadok[i])

    worksheet.conditional_format(1 + 3*j,0,1 + 3*j, poc_buniek, {'type':'2_color_scale',
                                                                 'min_color':'#000000',
                                                                 'max_color':'#00ff00',
                                                                 'min_value':0,
                                                                 'max_value':255})

#modre policka

for j in range(poc_riad):
    riadok = [int(bunka) for bunka in b[j]]
    for i in range(len(riadok)):
        worksheet.write(2 + 3*j, i, riadok[i])

    worksheet.conditional_format(2 + 3*j,0,2 + 3*j, poc_buniek, {'type':'2_color_scale',
                                                                 'min_color':'#000000',
                                                                 'max_color':'#0000ff',
                                                                 'min_value':0,
                                                                 'max_value':255})

workbook.close()
