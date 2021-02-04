
# CHART CATEGORIES
chartLabel = df_Categories_count['Categories'].tolist()
chartValue = df_Categories_count['Total'].tolist()

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig2 = plt.figure()
plt.pie(chartValue, labels=chartLabel, colors=colors, autopct='%1.1f%%', shadow=False, startangle=90)
plt.axis('equal')
plt.title("Categories", pad=20, fontsize=15)

fig2.savefig('C:/Users/Georges/Downloads/myplot2.png', dpi=75)
# plt.show()
plt.clf()

im = Image.open('C:/Users/Georges/Downloads/myplot2.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save('C:/Users/Georges/Downloads/myplot2.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image('C:/Users/Georges/Downloads/myplot2.png')
img.anchor = 'E6'

workbook['Categories'].add_image(img)
workbook.save(outputExcelFile)