import nex
import os

# specify your data directory and file extensions here
# note that we need to use double backslashes (C:\\Data, not C:\Data)
# dataDirectory = 'D:\\'
dataDirectory = 'E:\\'
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
    if doc:
        # apply autocorrelograms template
        # here we assume that we already saved the template with the name Autocorrelograms.ntp
        # some data files may not have any neurons or events
        # for these files, autocorrelograms cannot be calculated
        # we will use try/except to print the error and continue the script
        try:
            try:
                doc['stimstart'] = nex.FirstAfter(doc['Chan129a'], doc['Chan129a'], -20, 20)
                nex.DeselectAll(doc)
                nex.SelectAllNeurons(doc)
                nex.Deselect(doc, doc["Chan129a"])
            except:
                try:
                    doc['stimstart'] = nex.FirstAfter(doc['ainp1a'], doc['ainp1a'], -20, 20)
                    nex.DeselectAll(doc)
                    nex.SelectAllNeurons(doc)
                    nex.Deselect(doc, doc["ainp1a"])
                except:
                    try:
                        doc['stimstart'] = nex.FirstAfter(doc['ainp1U'], doc['ainp1U'], -20, 20)
                        nex.DeselectAll(doc)
                        nex.SelectAllNeurons(doc)
                        nex.Deselect(doc, doc["ainp1U"])
                    except Exception as ex:
                        print (ex)
                        break
            nex.ApplyTemplate(doc, "MUAPEH2")
            nex.SendResultsToExcel(doc, savefile+"_MUA.xlsx", "Nex", 0, "A1", 1, 0)
            nex.CloseExcelFile(savefile+"_MUA.xlsx")
            nex.SaveGraphics(doc,savefile+"_MUA.PNG",1)
        except Exception as ex:
            print(ex)
        nex.Sleep(1000)
        try:
            nex.ApplyTemplate(doc, "FirstPulses")
            nex.SendResultsToExcel(doc, savefile+"_Pulse.xlsx", "Nex", 0, "A1", 1, 0)
            nex.CloseExcelFile(savefile+"_Pulse.xlsx")
            nex.SaveGraphics(doc,savefile+"_Pulse.PNG",1)
        except Exception as ex:
            print(ex)
            break
        nex.Sleep(1000)        
        try:
            nex.DeselectAll(doc)
            var = 2
            while var <18:
                nex.SelectVar(doc,var,"continuous")
                var += 1
            nex.ApplyTemplate(doc, "LFPTIMEPSD_Final")
            nex.SendResultsToExcel(doc, savefile+"_PSD.xlsx", "Nex", 0, "A1", 1, 0)
            nex.CloseExcelFile(savefile+"_PSD.xlsx")
            nex.SaveGraphics(doc,savefile+"_PSD.PNG",1)
        except Exception as ex:
            print(ex)
        nex.Sleep(1000)     
        try:
            nex.ApplyTemplate(doc, "LFPSPEC_Final")
            nex.SendResultsToExcel(doc, savefile+"_LFP.xlsx", "Nex", 0, "A1", 1, 0)
            nex.CloseExcelFile(savefile+"_LFP.xlsx")
            nex.SaveGraphics(doc,savefile+"_LFP.PNG",1)
        except Exception as ex:
            print(ex)

        # delay the script for 500 milliseconds
        nex.Sleep(1000)
        
        # break

        # close the document
        nex.CloseDocument(doc)
        
print('Done!')


