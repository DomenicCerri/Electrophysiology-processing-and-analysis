import nex
import os

# specify your data directory and file extensions here
# note that we need to use double backslashes (C:\\Data, not C:\Data)
# dataDirectory = 'D:\\MUA_LFP\\PfT\\eYFP'
dataDirectory = 'E:\\MUA_LFP\\PfT\\eYFP'
block = 10
fileExtensions = ['.nex5']

def ListFiles(dataFilesDir, extensions):
    # list files with specified extensions in a specified directory 
    filePaths = []
    for f in os.listdir(dataFilesDir):
        path = os.path.join(dataFilesDir,f)
        if os.path.splitext(path)[1] in extensions:
            filePaths.append(path)
    return filePaths

filePaths = ListFiles(dataDirectory, fileExtensions)

# print the number of files
print('found {} files'.format(len(filePaths)))

# loop over files
for fileName in filePaths:
    
    # print the file name
    print(fileName)
   
    # open the file
    doc = nex.OpenDocument(fileName)
    savefile = nex.GetDocPath(doc)
    # if file was opened successfully
    blockshift = 0-(block+0.025)
    if doc:
        # apply autocorrelograms template
        # here we assume that we already saved the template with the name Autocorrelograms.ntp
        # some data files may not have any neurons or events
        # for these files, autocorrelograms cannot be calculated
        # we will use try/except to print the error and continue the script
        try:
            nex.Rename(doc, doc['ainp1a'], 'Chan129a')
        except:                   
            try:
                nex.Rename(doc, doc['ainp1U'], 'Chan129a')
            except:
                pass        
        try:
            doc['Chan129aBL'] = nex.Shift(doc['Chan129a'], blockshift)
            nex.DeselectAll(doc)
            nex.SelectAllNeurons(doc)
            nex.Deselect(doc, doc['Chan129a'])
            nex.Deselect(doc, doc['Chan129aBL'])
        except Exception as ex:
            print (ex)
            break
            
        try:
            nex.ApplyTemplate(doc, "BLPulses_50ms")
            nex.SendResultsToExcel(doc, savefile+"_BLPulses.xlsx", "Nex", 0, "A1", 1, 0)
            nex.CloseExcelFile(savefile+"_BLPulses.xlsx")
            nex.SaveGraphics(doc,savefile+"_BLPulses.PNG",1)
        except Exception as ex:
            print(ex)
            break       

            
        # delay the script for 500 milliseconds
        nex.Sleep(1000)
        
        # break

        # close the document
        nex.CloseDocument(doc)
        
print('Done!')
