clear all
close all

%% ~~~~~~~~~~~ Download excel sheet

Fixationtime = readtable("Event Statistics - fixation.xlsx");
% extract the relevant columns
fix = [Fixationtime.Participant, Fixationtime.EventStartTrialTime_ms_, Fixationtime.EventEndTrialTime_ms_, Fixationtime.EventDuration_ms_];
participant_namelist = ["privacy_01" "privacy_02" "privacy_03" "privacy_04" "privacy_05" "privacy_06" "privacy_07" "privacy_08" "privacy_09" "privacy_10" "privacy_11" "privacy_12" "privacy_13" "privacy_14" "privacy_15" "privacy_16" "privacy_17" "privacy_18" "privacy_19" "gprivacy_20" "privacy_21" "privacy_22" "privacy_23" "privacy_24"];
file = dir('/Users/muku/Documents/MATLAB folder/Ludvik/excel/video*.xlsx');
path2data = '/Users/muku/Documents/MATLAB folder/Ludvik/excel';


% ~~~~~~~~~~~~~~ each subjects loop start
for i = 1:numel(file) % number of the subject
	fn = [path2data '/' file(i).name]; % subjects path
    data = readtable(fn) % read excel sheet of the subject

    % ~~~~~~~~~~~ extract the trial number (Var1), release time (Var5), releasetime-2000 (Var6)
    a = table2array(data(:,["Var1","Var3","Var5","Var6"]))
    time = a(~isnan(a)) % remove NaN
    timepoint = [time(1:35,1),time(36:70,1),time(71:105,1),time(106:140,1)] %transforming 1 column to 3 column

    % ~~~~~~~~~~~ get index corresponding to the subjects name and relevant time
    nameindex = find(fix(:,1)==participant_namelist(i)) 
    d = str2double(fix(nameindex,2:4)) % remove the column including the subjects name
    subject_data = round(d) %remove the decimal point

    % ~~~~~~~~~~~ compare release time with eye trackers data
    for trialcount = 1:35 % number of trial

        timeindex_before_1 = subject_data(:,1) >= timepoint(trialcount,4)
        timeindex_release_1 = subject_data(:,1) <= timepoint(trialcount,3)
        targetindex_1 = timeindex_before_1 == timeindex_release_1 % index including both before and release
        
        timeindex_before_2 = subject_data(:,2) >= timepoint(trialcount,4)
        timeindex_release_2 = subject_data(:,2) <= timepoint(trialcount,3)
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
    subject_table = array2table(timepoint,"VariableNames",["Trial" "Movement_initiation" "Release_Time" "Release_2000"])
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
    % transform array to table
    sum_table = array2table(datasum,"VariableNames",["Sum of Duration"])
    lefttime_table = array2table(data_lefttime,"VariableNames",["Onset___Release2000 "])
    overtime_table = array2table(data_overtime,"VariableNames",["Offset___Release"])
    
    
    % ~~~~~~~~~~~~ extract Quiet Eye (QE)
    for k = 1:35

        QE_onset = cell2mat(dataTable.Onset(k));
        QE_offset = cell2mat(dataTable.Offset(k));
        QE_duration = cell2mat(dataTable.Duration(k));
        % no QE
        if isempty(QE_onset) || QE_onset(1) > subject_table.Movement_initiation(k);
            QE_time_onset(k,:) = [NaN];
            QE_time_offset(k,:) = [NaN];
            QE_time_duration(k,:) = [NaN];
        % QE 
        else
            indexQE = cell2mat(dataTable.Onset(k)) < subject_table.Movement_initiation(k);
            index = find(indexQE);
            lastTrueIndex = index(end);
            QE_time_onset(k,:) = QE_onset(lastTrueIndex);
            QE_time_offset(k,:) = QE_offset(lastTrueIndex);
        end

    end
    % transform array to table
    QEtotal = QE_time_offset - QE_time_onset;
    QEtotal_table = array2table(QEtotal,"VariableNames",["QE_duration(onset-offset)"]);
    QEonset_table = array2table(QE_time_onset,"VariableNames",["QE_onset"]);
    QEoffset_table = array2table(QE_time_offset,"VariableNames",["QE_offset"]);
    

    % ~~~~~~~~~~~ output excel sheet of subject(i)
    dataTable = [subject_table dataTable lefttime_table overtime_table sum_table Num_fix QEonset_table QEoffset_table QEtotal_table] 
    filename = 'MATLABfixation_data.xlsx'; % Specify the filename for the Excel file
    writetable(dataTable, filename, 'sheet', participant_namelist(i)); % output excel sheet


    % ~~~~~~~~~~~ clear all aside from these variables
    clearvars -except Fixationtime fix participant_namelist file path2data;


end
