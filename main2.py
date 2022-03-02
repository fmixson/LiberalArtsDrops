import pandas as pd
import openpyxl

class ConstructDataframe:
    def __init__(self, file):
        self.file = file

    def dataframe(self):
        pd.set_option('display.max_columns', None)
        # laDrops_df = pd.DataFrame(columns=headers)
        # file = 'D:\Work\Data\LA_Division\WeeklyEnrollmentSheets\Liberal Arts.csv'
        Enrollment_df = pd.read_csv(self.file)
        lectureEnrollment = Enrollment_df['Instruction Mode Description'] != 'Laboratory'
        lectureEnrollment_df = Enrollment_df[lectureEnrollment]
        return lectureEnrollment_df

class CourseDataFrame:

    def __init__(self, course, dataframe):
         self.course = course
         self.dataframe = dataframe

    def courseFilter(self):

        Drops = self.dataframe['Course'] == self.course
        Drops_df = self.dataframe[Drops]
        Drops_df = Drops_df.reset_index(drop=True)
        Drops_df.sort_values(by=['Employee ID', 'Enrollment Add Date'], inplace=True)
        Drops_df = Drops_df.reset_index(drop=True)
        Drops_df['Enrollment Drop Date'] = Drops_df['Enrollment Drop Date'].fillna(0)
        return Drops_df, self.course

    def divisionFilter(self):

        Drops_df = self.dataframe
        Drops_df.sort_values(by=['Employee ID', 'Enrollment Add Date'], inplace=True)
        Drops_df = Drops_df.reset_index(drop=True)
        Drops_df['Enrollment Drop Date'] = Drops_df['Enrollment Drop Date'].fillna(0)
        return Drops_df


class DataFrameCleanup:
    def __init__(self, Drops_df, course):
        self.course = course
        self.Drops_df = Drops_df
        self.multiplesList = []
        self.dropIndexList = []
        self.Drops_df.to_excel('Debugging.xlsx')

    def ReEnrolledList(self):
        for i in range(len(self.Drops_df)-1):
            # print(self.Drops_df)

            if self.Drops_df.loc[i, 'Employee ID'] == self.Drops_df.loc[i+1, 'Employee ID']:
                self.Drops_df.drop(index=i)
                self.multiplesList.append(self.Drops_df.loc[i, 'Employee ID'])
        print(self.multiplesList)
        return self.multiplesList


    def RemoveReEnrolled(self):
        for i in range((len(self.Drops_df)-1),-1,-1):
            for number in self.multiplesList:
                if number == self.Drops_df.loc[i, 'Employee ID']:
                    if self.Drops_df.loc[i, 'Enrollment Drop Date'] != 0:
                        if i not in self.dropIndexList:
                            self.dropIndexList.append(i)
        print(self.dropIndexList)
        self.Drops_df = self.Drops_df.drop(self.dropIndexList)
        dropZeros_df = self.Drops_df.loc[self.Drops_df['Enrollment Drop Date'] != 0]
        dropZeros_df['Enrollment Drop Date'] = pd.to_datetime(dropZeros_df['Enrollment Drop Date'])
        Drops = dropZeros_df['Enrollment Drop Date'] >= 'Jan 10, 2022'
        Drops_df = dropZeros_df[Drops]
        Drops_df = Drops_df.reset_index(drop=True)
        dropDate_series = Drops_df.groupby(['Enrollment Drop Date'], sort=True).size()
        Drops_df.to_excel(self.course + ' Drops.xlsx')
        dropDate_series.to_excel('DropDateSeries.xlsx')
        return Drops_df

class EOPSDrops:
    def __init__(self, Drops_df, course):
        self.Drops_df = Drops_df
        self.course = course

    def compareLists(self):
        EOPSstudents = pd.read_csv('EOPS.csv')
        for j in range(len(EOPSstudents)):
            for i in range(len(self.Drops_df) - 1):
                if EOPSstudents.loc[j, 'ID'] == self.Drops_df.loc[i, 'Employee ID']:
                    self.Drops_df.loc[i, 'EOPS'] = 'EOPS'
        EOPS = self.Drops_df.loc[self.Drops_df['EOPS'] =='EOPS']

        EOPS.to_excel(self.course + ' EOPSDrops.xlsx')

class PlotENGL100Drops:

    def __init__(self):

# update_df = update_df.loc[update_df['Enrollment Drop Date'] != 0]
# update_df = update_df.reset_index()
# # update_df.index.name = 'y'
#
# dropDate = update_df.groupby(['Enrollment Drop Date'], sort=True).size()
# print(dropDate)
def courseDropFunction(course):
    c = ConstructDataframe('D:\Work\Data\LA_Division\WeeklyEnrollmentSheets\Liberal Arts.csv')
    dataframe = c.dataframe()
    cd = CourseDataFrame(course=course, dataframe=dataframe)
    drops, course = cd.courseFilter()
    dd = DataFrameCleanup(Drops_df=drops, course=course)
    dd.ReEnrolledList()
    Drops_df = dd.RemoveReEnrolled()
    e = EOPSDrops(Drops_df=Drops_df, course=course)
    e.compareLists()

def divisionDropFunction():
    c=ConstructDataframe('D:\Work\Data\LA_Division\WeeklyEnrollmentSheets\Liberal Arts.csv')
    dataframe = c.dataframe()
    cd = CourseDataFrame(course='', dataframe=dataframe)
    drops = cd.divisionFilter()
    dd = DataFrameCleanup(Drops_df=drops, course='')
    dd.ReEnrolledList()
    Drops_df = dd.RemoveReEnrolled()



# courseDropFunction('English 100')
# courseDropFunction('English 103')
divisionDropFunction()
# print(dataframe)
# print(drops)




# laEnrollment_df.dropna(subset=['Enrollment Drop Date'], inplace=True)
# dropDate.to_excel('DropDate.xlsx')
# update_df.to_excel('100DropsData.xlsx')
# Engl100Drops_df.to_excel('Eng100sheet.xlsx')

