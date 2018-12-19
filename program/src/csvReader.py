import pandas as pd

class csvReader():

    def __init__(self,filename):
        self.__filename=filename
        self.__headNames=['code','name', 'department', 'lections', 'labs', 'practice', 'homework', 'examination']

    def csvReadInList(self):
        df=[]
        try:
            print("Reading csv file:", self.__filename)
            df=pd.read_csv(self.__filename, sep=';', header=0).values
        except FileNotFoundError:
            print("ERROR: Csv file not found")
            exit()
        return df
