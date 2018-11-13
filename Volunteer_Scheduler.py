import xlwings as xw
from pandas import DataFrame

class Volunteer():
    id = -1
    consecutiveWorkday = 0
    totalDaysOff = 0
    schedule = []

    def __init__(self, id, consecutiveWorkingDay,totalDaysOff,schedule):
        self.id = id
        self.consecutiveWorkday = consecutiveWorkingDay
        self.totalDaysOff = totalDaysOff
        self.schedule = schedule

    def clone(self):
        return Volunteer(self.id,self.consecutiveWorkday,self.totalDaysOff,self.schedule)

def addWorkday(volunteer, site):
    volunteer.schedule.append(site)
    volunteer.consecutiveWorkday = volunteer.consecutiveWorkday + 1
    return volunteer

def createVolunteerList(numVolunteer):
    volunteerList = []
    for i in range(numVolunteer):
        current = Volunteer(-1,0,0,[])
        current.id = i
        volunteerList.append(current)
    return volunteerList

def removeWeekend(workingList,remainingDaysOffPerDay):
        if len(workingList[0].schedule) != 0:
            for volunteer in workingList:
                if remainingDaysOffPerDay > 0:
                    if volunteer.schedule[-1] == 0:
                        if len(volunteer.schedule) == 1:
                            workingList.remove(volunteer)
                            remainingDaysOffPerDay = remainingDaysOffPerDay - 1
                        elif volunteer.schedule[-2] != 0:
                            workingList.remove(volunteer)
                            remainingDaysOffPerDay = remainingDaysOffPerDay - 1

def removeOverworked(workingList,remainingDaysOffPerDay, maxDaysWorking):
    for volunteer in workingList:
        if remainingDaysOffPerDay > 0 and volunteer.consecutiveWorkday >= maxDaysWorking:
            workingList.remove(volunteer)
            remainingDaysOffPerDay = remainingDaysOffPerDay - 1

def printVolunteerIds(volunteerList):
    for volunteer in volunteerList: print(volunteer.id)

def sortByTotalDaysOff(volunteerList):
    for i in range(len(volunteerList)):
        for j in range(len(volunteerList)):
            if volunteerList[i].totalDaysOff < volunteerList[j].totalDaysOff:
                temp = volunteerList[i]
                volunteerList[i] = volunteerList[j]
                volunteerList[j] = temp
            elif volunteerList[i].totalDaysOff == volunteerList[j].totalDaysOff and volunteerList[i].consecutiveWorkday < volunteerList[j].consecutiveWorkday:
                temp = volunteerList[i]
                volunteerList[i] = volunteerList[j]
                volunteerList[j] = temp

def numTimesWorkedSite(volunteer,site):
    numTimes = 0
    for currentSite in volunteer.schedule:
        if site == currentSite:
            numTimes = numTimes + 1
    return numTimes

def sortByTimesWorked(volunteerList,site):
    site = site + 1
    for i in range(len(volunteerList)):
        for j in range(len(volunteerList)):
            if numTimesWorkedSite(volunteerList[i],site) < numTimesWorkedSite(volunteerList[j],site):
                temp = volunteerList[i]
                volunteerList[i] = volunteerList[j]
                volunteerList[j] = temp

def distributeRemainingDaysOff(workingList,remainingDaysOffPerDay,minDaysBetweenWeekends):
    while(remainingDaysOffPerDay > 1):
        for volunteer in workingList:
            if volunteer.consecutiveWorkday >= minDaysBetweenWeekends and remainingDaysOffPerDay > 0:
                workingList.remove(volunteer)
                remainingDaysOffPerDay = remainingDaysOffPerDay - 1
        minDaysBetweenWeekends = minDaysBetweenWeekends - 1
    return workingList

def mergeVolunteerToList(volunteerList,volunteer):
    for i in range(len(volunteerList)):
        if volunteerList[i].id == volunteer.id:
            volunteerList[i] = volunteer
    return volunteerList

def getVolunteerById(volunteerList,id):
    for volunteer in volunteerList:
        if volunteer.id == id: return volunteer

def volunteerListToSchedule(volunteerList, numDays):
    schedule = DataFrame()
    for day in range(numDays):
        row = 0
        for volunteer in volunteerList:
            schedule.set_value(row, day, volunteer.schedule[day])
            row = row + 1
    return schedule

def main():
    #initialize inputs
    sht = xw.Book.caller().sheets[0]
    volunteer_schedule = sht.range("r_volunteer_schedule")
    numVolunteer = int(sht.range("n_volunteers").value)
    numDays = int(sht.range("n_days").value)
    numSites = sht.range("n_num_sites").value
    maxDaysWorking = sht.range("N_max_work_week").value
    minDaysBetweenWeekends = sht.range("N_min_work_week").value

    #create a list of empty volunteers
    volunteerList = createVolunteerList(numVolunteer)

    # Num of days off that can be assigned every day.
    daysOffPerDay = numVolunteer - numSites

    #Iterate through each day and assign sites and days off depending on volunteers past experience.
    for day in range(numDays):

        workingList = []
        for volunteer in volunteerList: workingList.append(volunteer.clone())
        remainingDaysOffPerDay = daysOffPerDay + len(workingList) - numVolunteer
        removeOverworked(workingList,remainingDaysOffPerDay,maxDaysWorking)

        remainingDaysOffPerDay = daysOffPerDay + len(workingList) - numVolunteer

        removeWeekend(workingList,remainingDaysOffPerDay)
        remainingDaysOffPerDay = daysOffPerDay + len(workingList) - numVolunteer

        sortByTotalDaysOff(workingList)
        distributeRemainingDaysOff(workingList,remainingDaysOffPerDay,minDaysBetweenWeekends)

        for site in range(int(numSites)):
            sortByTimesWorked(workingList,site)
            workingList[0] = addWorkday(workingList[0],site+1)
            volunteerList = mergeVolunteerToList(volunteerList,workingList[0])
            del workingList[0]
        for volunteer in volunteerList:
            if len(volunteer.schedule) < day+1:
                volunteer.schedule.append(0)
                volunteer.totalDaysOff = volunteer.totalDaysOff + 1
                volunteer.consecutiveWorkday = 0


    sht.range("r_volunteer_schedule").options(index=False, header=False).value = volunteerListToSchedule(volunteerList,numDays)
