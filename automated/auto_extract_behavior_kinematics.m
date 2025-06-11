% Prompt user to select the parent directory
parentDir = uigetdir([], 'Select the Parent Directory'); 
if parentDir == 0
    disp('No directory selected. Exiting.');
    return;
end

% Get all subfolders
allSubfolders = dir(parentDir);
allSubfolders = allSubfolders([allSubfolders.isdir]);
allSubfolders = allSubfolders(~ismember({allSubfolders.name}, {'.', '..'}));

% Define output file paths
exploreFile = fullfile(parentDir, 'Compiled_ExplorationBouts.xlsx');
kinematicsFile = fullfile(parentDir, 'Compiled_Kinematics.xlsx');

% Loop through each subfolder
for i = 1:length(allSubfolders)
    subfolderName = allSubfolders(i).name;
    subfolderPath = fullfile(parentDir, subfolderName);

    %% === EXPLORATION BOUTS ===
    behaviorFile = fullfile(subfolderPath, 'Behavior.mat');
    T_explore = table();

    if exist(behaviorFile, 'file') == 2
        S = load(behaviorFile);
        if isfield(S, 'Behavior') && isfield(S.Behavior, 'Exploration') && isfield(S.Behavior.Exploration, 'Bouts')
            A_data = S.Behavior.Exploration.Bouts;
            if isnumeric(A_data) && size(A_data,2) == 2
                T_explore = array2table(A_data, 'VariableNames', {'StartFrame', 'EndFrame'});

                % Try to add frameRate
                if endsWith(subfolderName, '_Obj2')
                    paramsFile = fullfile(subfolderPath, 'Params.mat');
                    if exist(paramsFile, 'file') == 2
                        try
                            P = load(paramsFile);
                            if isfield(P, 'Params') && isfield(P.Params, 'Video') && isfield(P.Params.Video, 'frameRate')
                                frameRate = P.Params.Video.frameRate;
                                if isnumeric(frameRate) && isscalar(frameRate)
                                    T_explore.frameRate = repmat(frameRate, height(T_explore), 1);
                                end
                            end
                        catch
                            warning('Could not load frameRate from Params.mat in %s', subfolderName);
                        end
                    end
                end
            else
                warning('Invalid Bouts format in %s', subfolderName);
            end
        end
    end

    %% === KINEMATICS (only for _Obj1 folders) ===
    T_kinematics = table();
    isObj1 = contains(subfolderName, '_Obj1');
    if isObj1
        metricsFile = fullfile(subfolderPath, 'Metrics.mat');
        if exist(metricsFile, 'file') == 2
            try
                load(metricsFile);
                A = Metrics.Movement.DistanceTraveled(:);  
                B = Metrics.Velocity.Head(:);
                C = Metrics.Acceleration.Head(:);
                D = Metrics.Velocity.MidBack(:);
                E = Metrics.Acceleration.MidBack(:);
                T_kinematics = table(A, B, C, D, E, ...
                    'VariableNames', {'DistanceTraveled', 'VelocityHead', 'AccelerationHead', 'VelocityMidBack', 'AccelerationMidBack'});
            catch
                warning('Could not extract Metrics in %s', subfolderName);
            end
        end
    end

    %% === CLEAN SHEET NAMES ===
    sheetNameExplore = regexprep(subfolderName, '[:\\/*?\[\]]', '_');
    if length(sheetNameExplore) > 31
        sheetNameExplore = sheetNameExplore(1:31);
    end

    sheetNameKinematics = sheetNameExplore;
    if isObj1
        sheetNameKinematics = strrep(sheetNameExplore, '_Obj1', '');
        if length(sheetNameKinematics) > 31
            sheetNameKinematics = sheetNameKinematics(1:31);
        end
    end

    %% === WRITE TO EXCEL ===
    if ~isempty(T_explore)
        try
            writetable(T_explore, exploreFile, 'Sheet', sheetNameExplore);
            fprintf('✅ Wrote exploration data for %s\n', subfolderName);
        catch ME
            warning('❌ Failed to write exploration sheet %s: %s', sheetNameExplore, ME.message);
        end
    end

    if ~isempty(T_kinematics)
        try
            writetable(T_kinematics, kinematicsFile, 'Sheet', sheetNameKinematics);
            fprintf('✅ Wrote kinematics data for %s\n', subfolderName);
        catch ME
            warning('❌ Failed to write kinematics sheet %s: %s', sheetNameKinematics, ME.message);
        end
    end
end

disp(['✅ Exploration data saved to: ', exploreFile]);
disp(['✅ Kinematics data saved to: ', kinematicsFile]);