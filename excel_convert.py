from xlrd import open_workbook
import sys, getopt

def main(argv):
    inputfile = ''
    outputfile = ''
    try:
        opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print 'excel_convert.py -i <inputfile> -o <outputfile>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print 'excel_convert.py -i <inputfile> -o <outputfile>'
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg

    wb = open_workbook(inputfile)

    outFile = open(outputfile,'w')
    outFile.write('"sep=;"\n')
    # Or grab the first sheet by index 
    #  (sheets are zero-indexed)
    #
    xl_sheet = wb.sheet_by_index(0)
    print ('Sheet name: %s' % xl_sheet.name)


    for i in xrange(xl_sheet.nrows):
        cell = xl_sheet.cell(i, 0)
        if cell.value.startswith("Measurement count:"):
            print cell.value
            #outFile.write('%s\n'%cell.value)

            xOffset = 1
            yOffset = i+2
            xSize = 13
            ySize = yOffset + 8
            cells = []
            cells.append(cell.value)
            for x in xrange(xOffset, xSize):
                for y in xrange(yOffset, ySize):
                    cell = xl_sheet.cell(y, x)
                    cells.append(cell.value)
            print cells  
            outFile.write('%s\n' %";".join(map(str, cells)))


    outFile.close()    
    
    

if __name__ == "__main__":
   main(sys.argv[1:])