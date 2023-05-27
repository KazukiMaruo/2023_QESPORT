clear all
close all

%% ~~~~~~~~~~~ Download excel sheet

Fixationtime = readtable("Event Statistics - fixation.xlsx");
% extract the relevant columns
fix = [Fixationtime.Participant, Fixationtime.EventStartTrialTime_ms_, Fixationtime.EventEndTrialTime_ms_, Fixationtime.EventDuration_ms_];
participant_namelist = ["subject_01" "subject_02" "subject_03" "subject_04" "subject_05" "subject_06" "subject_07" "subject_08" "subject_09" "subject_10" "subject_11" "subject_12"];

file = dir('/Users/muku/Documents/MATLAB folder/Ludvik/excel/video*.xlsx');
path2data = '/Users/muku/Documents/MATLAB folder/Ludvik/excel';


% ~~~~~~~~~~~~~~ each subjects loop start
for i = 1:numel(file) % number of the subject
	fn = [path2data '/' file(i).name]; %subjects path
    % read excel sheet of the subject
    data = readtable(fn)


    % ~~~~~~~~~~~ extract the trial number (Var1), release time (Var5), releasetime-2000 (Var6)
    a = table2array(data(:,["Var1","Var5","Var6"]))
    time = a(~isnan(a)) % remove NaN
    timepoint = [time(1:35,1),time(36:70,1),time(71:105,1)] %transforming 1 column to 3 column


    % ~~~~~~~~~~~ get index corresponding to the subject's name and relevant time
    nameindex = find(fix(:,1)==participant_namelist(i)) 
    d = str2double(fix(nameindex,2:4)) % remove the column including the subjects name
    subject_data = round(d) %remove the decimal point


    % ~~~~~~~~~~~ compare release time with eye tracker's data
    for trialcount = 1:35 % number of trial
        timeindex_before_1 = subject_data(:,1) >= timepoint(trialcount,3)
        timeindex_release_1 = subject_data(:,1) <= timepoint(trialcount,2)
        targetindex_1 = timeindex_before_1 == timeindex_release_1 % index including both before and release
        
        timeindex_before_2 = subject_data(:,2) >= timepoint(trialcount,3)
        timeindex_release_2 = subject_data(:,2) <= timepoint(trialcount,2)
        targetindex_2 = timeindex_before_2 == timeindex_release_2 % index including both before and release
        
        % extract the corresponding time with the subejcts fixation
        if  any(targetindex_1) | any(targetindex_2) % if either index is true
            Fixation_list{trialcount,1} = subject_data(targetindex_1 | targetindex_2,1) 
            Fixation_list{trialcount,2} = subject_data(targetindex_1 | targetindex_2,2)
            Fixation_list{trialcount,3} = subject_data(targetindex_1 | targetindex_2,3)
            count_fixation(trialcount,:) = numel(subject_data(targetindex_1 | targetindex_2,1))
        else % if both are false, write nan
            Fixation_list{trialcount,1} = []
            Fixation_list{trialcount,2} = []
            Fixation_list{trialcount,3} = []
            count_fixation(trialcount,:) = [NaN]
        end 

    end


    % ~~~~~~~~~~~ Create table to output excel sheet
    dataTable = cell2table(Fixation_list,"VariableNames",["Onset" "Offset" "Duration" ])
    subject_table = array2table(timepoint,"VariableNames",["Trial" "Release_Time" "Release_2000"])
    Num_fix = array2table(count_fixation,"VariableNames",["Num of Fixation"])

    % calculate the sum of duration and 
    % time outside between release and release -2000
    for s = 1:35
        datasum(s,:) = sum(cell2mat(dataTable.Duration(s)))
        lefttime = cell2mat(dataTable.Onset(s))
        if isempty(lefttime) || lefttime(1) > subject_table.Release_2000(s)
            data_lefttime(s,:) = [NaN]
        else
            data_lefttime(s,:) =  lefttime(1) - subject_table.Release_2000(s)
        end

        overtime = cell2mat(dataTable.Offset(s))
        if isempty(overtime) || overtime(end) < subject_table.Release_Time(s)
            data_overtime(s,:) = [NaN]
        else
            data_overtime(s,:) = overtime(end) - subject_table.Release_Time(s) 
        end

    end
    sum_table = array2table(datasum,"VariableNames",["Sum of Duration"])
    lefttime_table = array2table(data_lefttime,"VariableNames",["Onset___Release2000 "])
    overtime_table = array2table(data_overtime,"VariableNames",["Offset___Release"])


    % ~~~~~~~~~~~ output excel sheet
    dataTable = [subject_table dataTable lefttime_table overtime_table sum_table Num_fix] % 
    filename = 'fixation_lists.xlsx'; % Specify the filename for the Excel file
    writetable(dataTable, filename, 'sheet', participant_namelist(i)); % output excel sheet


    % ~~~~~~~~~~~ clear all aside from these variables
    clearvars -except Fixationtime fix participant_namelist file path2data;

end
