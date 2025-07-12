% =========================================================================
% LightTools Model Interrogation Script
% =========================================================================
%
% PURPOSE:
%   Connects to a running LightTools session and systematically extracts
%   all available data properties for a user-defined list of LightTools
%   objects (identified by their data access keys). It handles both scalar
%   and array-based data (like meshes), storing the results in a .mat file
%   for later analysis in MATLAB.
%
% AUTHOR: Atif Khan
% DATE: 13/7/2025
%
% =========================================================================

function InterrogateLightToolsModel()
    clear; clc; close all;

    disp('--- Script Start: LightTools Model Interrogation ---');

    % --- Configuration ---
    lightToolsVersion = '8.4.0'; % <<< CONFIRM this matches your LT version
    pid = 31912;                % <<< IMPORTANT: Update with the PID of YOUR RUNNING LT session
    saveDirectory = 'C:\Temp\LT_Interrogation_Output'; % <<< Directory for output
    
    % --- !!! DEFINE THE KEYS YOU WANT TO INTERROGATE HERE !!! ---
    keysToProcess = {
        'LENS_MANAGER[1].COMPONENTS[Components].SOLID[MyAutomatedSphere]', ...
        'LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[PlaneReceiver].FORWARD_SIM_FUNCTION[Forward_Simulation].ILLUMINANCE_MESH[Illuminance_Mesh]', ...
        'LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[PlaneReceiver].FORWARD_SIM_FUNCTION[Forward_Simulation].SPECTRAL_DISTRIBUTION[Spectral_Distribution_and_CRI]'
        % Add more keys as needed
    };
    % --- End Key Definition ---
    
    runIdentifier = datestr(now, 'yyyymmdd_HHMMSS');
    tempDumpFilePath = fullfile(saveDirectory, 'temp_keydump.txt');
    outputMatFileName = fullfile(saveDirectory, ['Model_Interrogation_Output_' runIdentifier '.mat']);
    outputCsvFileName = fullfile(saveDirectory, ['Model_Interrogation_Output_' runIdentifier '.csv']);

    if ~isfolder(saveDirectory), mkdir(saveDirectory); end
    disp(['Output will be saved in: ' saveDirectory]);

    lt = connectToLightTools(lightToolsVersion, pid);
    if isempty(lt), return; end % Exit if connection failed
    
    DBGET_SUCCESS_CODE = 0; % Adjust if your DbGet success code is different

    try
        allResultsData = {};     % Cell array for scalar values and status/placeholders
        arrayResultsData = struct(); % Struct to hold full array data

        for k_key = 1:length(keysToProcess)
            currentSourceKey = keysToProcess{k_key};
            fprintf('\nProcessing Key %d of %d: %s\n', k_key, length(keysToProcess), currentSourceKey);

            if isfile(tempDumpFilePath), delete(tempDumpFilePath); end
            dumpStatus = lt.DbKeyDump(currentSourceKey, tempDumpFilePath);
            if ~isfile(tempDumpFilePath)
                warning('DbKeyDump did not create temp file for key %s. Status was: %d. Skipping.', currentSourceKey, dumpStatus);
                continue;
            end

            propertiesToAttempt = parseDbKeyDumpFile(tempDumpFilePath);
            fprintf('  Parsed %d unique properties for this key.\n', size(propertiesToAttempt,1));
            
            % --- Get Live Values for this key's properties ---
            fprintf('  Getting live values...\n');
            for i_prop = 1:size(propertiesToAttempt,1)
                propName = propertiesToAttempt{i_prop, 1};
                dataType = propertiesToAttempt{i_prop, 2};
                currentRow = size(allResultsData, 1) + 1;
                allResultsData{currentRow, 1} = currentSourceKey;
                allResultsData{currentRow, 2} = propName;
                resultValueStorage = '<Error>'; % Default in case of issues

                try
                    if contains(dataType, '(ij)') % 2D Array Data (Meshes)
                        [resultValueStorage, arrayData] = getMeshData(lt, currentSourceKey, propName);
                        if isnumeric(arrayData) % If we got a numeric array back
                            safeKey = matlab.lang.makeValidName(currentSourceKey);
                            safePropName = matlab.lang.makeValidName(propName);
                            if ~isfield(arrayResultsData, safeKey), arrayResultsData.(safeKey) = struct(); end
                            arrayResultsData.(safeKey).(safePropName) = arrayData;
                        end
                    else % Scalar Data
                        [resultValueStorage] = getScalarData(lt, currentSourceKey, propName);
                    end
                catch ME_get
                    resultValueStorage = ['<MATLAB Error during Get: ' ME_get.message '>'];
                end
                allResultsData{currentRow, 3} = resultValueStorage;
            end
        end
        disp('--- Finished processing all keys ---');

        if isfile(tempDumpFilePath), delete(tempDumpFilePath); end
        
        writeResultsToFiles(outputCsvFileName, outputMatFileName, allResultsData, arrayResultsData, keysToProcess);

    catch ME
        disp(' ');
        fprintf(2, 'An error occurred in the main script: %s\n', ME.message);
        for k_stack = 1:length(ME.stack)
            fprintf(2, 'In file: %s, function: %s, at line: %d\n', ME.stack(k_stack).file, ME.stack(k_stack).name, ME.stack(k_stack).line);
        end
    end

    % --- Release COM Object ---
    if exist('lt','var') && iscom(lt), lt.delete; end
    if exist('asm','var'), clear asm; end
    disp('LightTools connection released.');
    disp('--- Script End ---');
end

% --- Helper Function for Connection ---
function lt = connectToLightTools(version, pid)
    lt = [];
    baseLtPath = 'C:\Program Files\Optical Research Associates\LightTools';
    ltcom64path = fullfile(baseLtPath, ['LightTools ' version], 'Utilities.NET', 'LTCOM64.dll');
    if ~isfile(ltcom64path), error('LTCOM64.dll not found at: %s.', ltcom64path); end
    try
        NET.addAssembly(ltcom64path);
        lt = LTCOM64.LTAPIx;
        lt.LTPID = pid;
        lt.UpdateLTPointer;
        lt.Message(['MATLAB connected to PID: ' num2str(pid)]);
        disp('Successfully connected to running LightTools session.');
    catch ME
        if ~isempty(lt) && isinterface(lt), lt.delete; end
        error('Failed to connect to LightTools PID %d. Error: %s', pid, ME.message);
    end
end

% --- Helper Function to Parse DbKeyDump File ---
function properties = parseDbKeyDumpFile(filePath)
    properties = {}; parsingStarted = false;
    processedPropNames = containers.Map('KeyType','char','ValueType','logical');
    fid = fopen(filePath, 'r');
    if fid == -1, warning('Could not open temp dump file for parsing.'); return; end
    tline = fgetl(fid);
    while ischar(tline)
        if contains(tline, 'Available functions for this data key'), parsingStarted = true; tline = fgetl(fid); tline = fgetl(fid); continue; end
        if parsingStarted
            if startsWith(strtrim(tline), 'Sub-Components') || isempty(strtrim(tline)), break; end
            tokens = regexp(tline, '^\s*(.+?)\s+(RW|RO)\s+([\w\(\)]+).*$', 'tokens', 'once');
            if ~isempty(tokens)
                propName = strtrim(tokens{1}); dataType = strtrim(tokens{3});
                if ~isKey(processedPropNames, propName)
                    properties{end+1, 1} = propName; properties{end, 2} = dataType;
                    processedPropNames(propName) = true;
                end
            end
        end
        tline = fgetl(fid);
    end
    fclose(fid);
end

% --- Helper Function to Get Mesh Data ---
function [resultStr, arrayData] = getMeshData(ltHandle, key, propName)
    arrayData = []; % Default empty
    [xDim, statusX] = ltHandle.DbGet(key, 'X_Dimension');
    [yDim, statusY] = ltHandle.DbGet(key, 'Y_Dimension');
    DB_SUCCESS = 0; % Assume 0 is success
    if statusX == DB_SUCCESS && statusY == DB_SUCCESS
        xDim = double(xDim); yDim = double(yDim);
        dummyData = zeros(1); cellFilter = propName;
        [statusMesh, meshDataOut] = ltHandle.GetMeshData(key, dummyData, cellFilter);
        if statusMesh ~= DB_SUCCESS % Or your API's success code
            if ~isa(meshDataOut, 'double'), arrayData = double(meshDataOut); else arrayData = meshDataOut; end
            dims = size(arrayData);
            resultStr = sprintf('<Array Data Stored Separately - Size [%d x %d]>', dims(1), dims(2));
        else
            [errStr, ~] = ltHandle.GetStatusString(statusMesh);
            resultStr = ['<GetMeshData Error: ' num2str(statusMesh) ' (' char(errStr) ')>'];
        end
    else
        resultStr = '<Error getting mesh dimensions>';
    end
end

% --- Helper Function to Get Scalar Data ---
function resultStr = getScalarData(ltHandle, key, propName)
    DB_SUCCESS = 0; % Assume 0 is success
    [value, status] = ltHandle.DbGet(key, propName);
    if status == DB_SUCCESS
        if isa(value, 'System.String'), resultStr = ['"' char(value) '"'];
        elseif isempty(value), resultStr = '<Empty>';
        elseif isnumeric(value), resultStr = num2str(value);
        elseif islogical(value), if value, resultStr = '"Yes"'; else resultStr = '"No"'; end
        elseif isinterface(value,'Interface'), resultStr = '<COM Object>';
        elseif isa(value, 'System.Object') && isprop(value,'VARIANT') && strcmp(class(value.VARIANT), 'System.__ComObject'), resultStr = '<COM Variant>';
        elseif iscell(value), resultStr = '<Cell Array>'; else resultStr = ['<Unsupported Type: ' class(value) '>']; end
    else
        [errStr, ~] = ltHandle.GetStatusString(status);
        resultStr = ['<DbGet Error: ' num2str(status) ' (' char(errStr) ')>'];
    end
end

% --- Helper Function to Write Output Files ---
function writeResultsToFiles(csvPath, matPath, scalarData, arrayData, processedKeys)
    disp(['Writing scalar results to CSV: ' csvPath]);
    try
        fid_csv = fopen(csvPath, 'w');
        if fid_csv == -1, error('Could not open CSV file for writing.'); end
        fprintf(fid_csv, 'ObjectKey,PropertyName,CurrentValue_Or_Status\n');
        for i = 1:size(scalarData, 1)
            csvValue = strrep(scalarData{i, 3}, '"', '""');
            fprintf(fid_csv, '"%s","%s","%s"\n', scalarData{i, 1}, scalarData{i, 2}, csvValue);
        end
        fclose(fid_csv);
        disp('Successfully wrote CSV file.');
    catch ME_csv
        if fid_csv ~= -1, fclose(fid_csv); end
        warning('Error writing to CSV file: %s', ME_csv.message);
    end

    disp(['Saving results data to MAT file: ' matPath]);
    try
        save(matPath, 'scalarData', 'arrayData', 'processedKeys');
        disp('Successfully saved results to .mat file.');
    catch ME_save
        warning('Error saving results to .mat file: %s', ME_save.message);
    end
end
