import win32com.client 
import datetime 
import pandas as pd 

class MsProject:
    

    def ms_project(file):
        try:
            msapp = win32com.client.Dispatch('MSPRoject.Application')
            msapp.Visible = 1 
        
            try:
                msapp.FileOpen(file) 
                msproj = msapp.ActiveProject
                msTaskList = msproj.Tasks 

                print(msproj.Name)  
            except Exception as e:
                print(e) 

            msapp.Quite()

        except Exception as e : 
            print("Error is : " ,  e)  
            