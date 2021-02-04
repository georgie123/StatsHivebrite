
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

pdf = PdfPages('C:/Users/Georges/Downloads/multipage.pdf')

fig1 = plt.figure()
plt.plot([0,1,2,3,4])
plt.close()
pdf.savefig(fig1)

fig2 = plt.figure()
plt.plot([0,2,4,6,8])
plt.close()
pdf.savefig(fig2)

pdf.close()