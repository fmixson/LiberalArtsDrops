import pandas as pd






# headers = ['Dept', 'Course', 'Session', 'Class', 'Start', 'End', 'Days', 'Room', 'Size', 'Max', 'Wait', 'Cap', 'Seats',
#            'WaitAv', 'Status', 'Instructor', 'Type', 'Hours', 'Books', 'Modality']
pd.set_option('display.max_columns', None)
# laDrops_df = pd.DataFrame(columns=headers)
file = 'D:\Work\Data\LA_Division\WeeklyEnrollmentSheets\Liberal_Arts.csv'
laEnrollment_df = pd.read_csv(file)
laLecture = laEnrollment_df['Instruction Mode Description'] != 'Laboratory'
lectureEnrollment_df = laEnrollment_df[laLecture]


# print(lectureEnrollment_df)
Engl100Drops = lectureEnrollment_df['Course'] == 'English 100'
Engl100Drops_df = lectureEnrollment_df[Engl100Drops]
Engl100Drops_df = Engl100Drops_df.reset_index()
Engl100Drops_df.sort_values(by=['Employee ID', 'Enrollment Add Date'], inplace=True)
Engl100Drops_df = Engl100Drops_df.reset_index()
Engl100Drops_df['Enrollment Drop Date'] = Engl100Drops_df['Enrollment Drop Date'].fillna(0)


multiplesList = []
dropIndexList = []
for i in range(len(Engl100Drops_df)-1):
    if Engl100Drops_df.loc[i, 'Employee ID'] == Engl100Drops_df.loc[i+1, 'Employee ID']:
        Engl100Drops_df.drop(index=i)
        multiplesList.append(Engl100Drops_df.loc[i, 'Employee ID'])


for i in range((len(Engl100Drops_df)-1),-1,-1):
    for number in multiplesList:
        if number == Engl100Drops_df.loc[i, 'Employee ID']:
            if Engl100Drops_df.loc[i, 'Enrollment Drop Date'] != 0:
                print(i, number, Engl100Drops_df.loc[i, 'Employee ID'])
                if i not in dropIndexList:
                    dropIndexList.append(i)
    print(dropIndexList)
    update_df = Engl100Drops_df.drop(dropIndexList)
update_df = update_df.loc[update_df['Enrollment Drop Date'] != 0]


# laEnrollment_df.dropna(subset=['Enrollment Drop Date'], inplace=True)

update_df.to_excel('data.xlsx')
Engl100Drops_df.to_excel('DropsEng100.xlsx')

