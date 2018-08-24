Function CalculateDueDate(StartDate As Date, TAT As Long, Optional IncludeDayZero As Boolean = False, Optional IncludeWeekends As Boolean = True) As Date
    
    'CalculateDueDate
    'Function for calculating the due date based on a start date and a given turnaround time (TAT)
    'Written By: Bryan O'Leary
    'Last Updated On: 05/07/2018 (dd/mm/yyyy)
     
    '''''''''''Variables
    'StartDate
    '   - The day to calculate the DueDate from
    'TAT (Turnaround Time)
    '   - The number of days in which the work is due
    'IncludeDayZero
    '   -   also the user to adjust determine whether the due date should be
    '       counted from the StartDate or from the following day and set rules accordingly,
    '       e.g. If a work item arrives too late at night and work can't begin until the next day
    '   -   The default value is False, so it assumes that the StartDate is counted
    'IncludeWeekends
    '   -   allows the user to specify whether or not to include weekends
    '   -   the default is True so weekends are included
    '   -   the function uses the Excel formula function Networkdays_Intl so the default weekend days can be altered if necessary
    
    Dim x As Long
    Dim CountWeekendDays As Long
    Dim WeekendDaysAccountedFor As Long
    Dim TempDueDate As Date
    
    If IncludeDayZero = False Then
        TAT = TAT - 1
    End If
    
    If IncludeWeekends Then
        CalculateDueDate = StartDate + TAT

    Else
        CountWeekendDays = ((StartDate + TAT) - StartDate) - Application.WorksheetFunction.NetworkDays_Intl(StartDate, StartDate + TAT, 1) + 1
        WeekendDaysAccountedFor = 0
        TempDueDate = (StartDate + TAT + CountWeekendDays)
        
        While WeekendDaysAccountedFor < CountWeekendDays
            TempDueDate = (StartDate + TAT + CountWeekendDays)
            WeekendDaysAccountedFor = CountWeekendDays
            CountWeekendDays = (TempDueDate - StartDate) - Application.WorksheetFunction.NetworkDays_Intl(StartDate, TempDueDate, 1) + 1
            
        Wend
        CalculateDueDate = TempDueDate
        
    End If
    
End Function