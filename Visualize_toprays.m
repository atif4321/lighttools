% =========================================================================
% MATLAB Script to Visualize Power Percentage Intervals of Ray Paths
% =========================================================================
%
% PURPOSE:
%   Connects to a running LightTools session, retrieves ray path data for a
%   specified receiver, and interactively filters the rays by source and
%   final surface. For user-defined power percentage intervals (e.g.,
%   [100,70],[70,30]), it isolates the corresponding rays, makes ONLY those
%   rays visible in the 3D view, captures a screenshot using the clipboard,
%   and saves the data of the visualized rays to a .mat file. This process
%   repeats for each defined interval.
%
% AUTHOR: Atif Khan
% DATE: 13/7/2025
%
% INSTRUCTIONS:
%   1. Launch LightTools and open your model.
%   2. Run a simulation that generates ray path data for your target receiver.
%   3. Update the 'Configuration' section below with your LightTools
%      version and the Process ID (PID) of the running LightTools session.
%   4. Run this script in MATLAB.
%   5. Follow the pop-up dialogs to define intervals and select filters.
%
% =========================================================================

function VisualizeTopRayPaths()
    clear; clc; close all;

    disp('--- Script Start: Visualizing Filtered Top Power Ray Paths ---');

    % --- Configuration ---
    lightToolsVersion = '8.4.0'; % <<< CONFIRM this matches your LT version
    pid = 31912;                % <<< IMPORTANT: Update with the PID of YOUR RUNNING LT session
    saveDirectory = 'C:\Temp\LightTools_Output\RayPathCaptures'; % <<< Directory for output
    outputImageBaseName = 'RayPathInterval';
    outputImageFormat = 'png'; % Output format for clipboard capture

    % --- Main Script Logic ---
    lt = []; % Initialize handle
    baseKeyForRayPaths = '';
    numRayPaths = 0;
    
    try
        % --- User Input ---
        [powerIntervalsCell, receiverName, wasCancelled] = getUserInput();
        if wasCancelled, return; end

        % --- Setup and Connection ---
        if ~isfolder(saveDirectory), mkdir(saveDirectory); end
        disp(['Output will be saved in: ' saveDirectory]);
        lt = connectToLightTools(lightToolsVersion, pid);
        if isempty(lt), return; end
        
        runIdentifier = datestr(now, 'yyyymmdd_HHMMSS');

        % --- Data Retrieval and Processing ---
        baseKeyForRayPaths = sprintf('LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[%s].FORWARD_SIM_FUNCTION[Forward_Simulation]', receiverName);
        [allRayData, numRayPaths] = retrieveRayPathData(lt, baseKeyForRayPaths);
        [sourceFilter, surfaceFilter, wasCancelled] = getInteractiveFilters(allRayData);
        if wasCancelled, return; end
        [sortedFilteredRayData, totalPowerOfFiltered] = filterAndSortRayData(allRayData, sourceFilter, surfaceFilter);

        if totalPowerOfFiltered <= 1e-12
            warning('Total power of filtered rays is effectively zero. No interval screenshots will be generated.');
            powerIntervalsCell = {};
        end

        % --- Main Loop: Process Each Power Percentage Interval ---
        for interval_idx = 1:length(powerIntervalsCell)
            currentInterval = powerIntervalsCell{interval_idx};
            upperPercent = currentInterval(1);
            lowerPercent = currentInterval(2);
            fprintf('\n--- Processing Power Interval: %.2f%% to %.2f%% ---\n', upperPercent, lowerPercent);

            % --- Select Rays for the Interval ---
            indicesToUpperSet = getRayIndicesForCumulativePercent(sortedFilteredRayData, totalPowerOfFiltered, upperPercent);
            indicesToLowerSet = getRayIndicesForCumulativePercent(sortedFilteredRayData, totalPowerOfFiltered, lowerPercent);
            if lowerPercent == 0
                indicesOfRaysInInterval = indicesToUpperSet;
            else
                indicesOfRaysInInterval = setdiff(indicesToUpperSet, indicesToLowerSet, 'stable');
            end
            fprintf('  Identified %d rays for this power interval.\n', length(indicesOfRaysInInterval));

            % --- Set Visibility and Capture Image ---
            setVisibility(lt, baseKeyForRayPaths, numRayPaths, indicesOfRaysInInterval);
            captureViaClipboard(lt, saveDirectory, outputImageBaseName, outputImageFormat, runIdentifier, currentInterval, sourceFilter, surfaceFilter);
            
            % --- Save Data for this interval's selected rays ---
            saveIntervalData(allRayData, indicesOfRaysInInterval, saveDirectory, outputImageBaseName, runIdentifier, currentInterval, sourceFilter, surfaceFilter);
            pause(1.0);
        end

    catch ME
        disp(' ');
        fprintf(2, 'An error occurred in the main script: %s\n', ME.message);
        for k_stack = 1:length(ME.stack)
            fprintf(2, 'In file: %s, function: %s, at line: %d\n', ME.stack(k_stack).file, ME.stack(k_stack).name, ME.stack(k_stack).line);
        end
    end

    % --- Final Cleanup ---
    restoreAllRayVisibility(lt, baseKeyForRayPaths, numRayPaths);
    releaseLightTools(lt);
    disp('--- Script End ---');
end


% =========================================================================
%                        HELPER FUNCTIONS
% =========================================================================

function [intervals, receiver, cancelled] = getUserInput()
    cancelled = false;
    intervals = {};
    receiver = '';
    
    prompt = {'Enter power percentage intervals (e.g., [[100,70],[70,30],[30,0]]):', ...
              'Enter Receiver Name:'};
    dlgTitle = 'Ray Path Visualization Setup';
    defInput = {'[[100,70],[70,30],[30,0]]', 'PlaneReceiver'};
    answer = inputdlg(prompt, dlgTitle, [1 100; 1 70], defInput);

    if isempty(answer), cancelled = true; disp('User cancelled input. Exiting.'); return; end
    powerIntervalsStr = strtrim(answer{1});
    receiver = strtrim(answer{2});

    try
        cleanedStr = regexprep(powerIntervalsStr, '^\[|\]$', '');
        intervalPairsStr = regexp(cleanedStr, '\]\s*,\s*\[', 'split');
        if isempty(intervalPairsStr{1}) && ~isempty(cleanedStr), intervalPairsStr = {cleanedStr}; end
        for k = 1:length(intervalPairsStr)
            pairNum = str2num(regexprep(intervalPairsStr{k}, '[\[\]]', ''));
            if isnumeric(pairNum) && numel(pairNum)==2, intervals{end+1} = pairNum; end
        end
        if isempty(intervals) && ~isempty(powerIntervalsStr), error('Parsing failed.'); end
        if ~all(cellfun(@(x) x(1)>=x(2) && all(x>=0) && all(x<=100), intervals)), error('Values must be 0-100 and upper>=lower.'); end
    catch ME
        error('Error parsing intervals: %s. Use format like [[100,70],[70,30]].', ME.message);
    end
    fprintf('Parsed %d power intervals.\n', length(intervals));
end

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

function [allRayData, numRayPaths] = retrieveRayPathData(lt, forwardSimKey)
    DBGET_SUCCESS_CODE = 0; % Adjust if your version's success code is different
    
    fprintf('  Getting total number of ray paths...\n');
    statusShow = lt.DbSet(forwardSimKey, 'ShowRayPaths', "Yes");
    if statusShow ~= DBGET_SUCCESS_CODE, warning('Could not enable ShowRayPaths. Status: %d', statusShow); end
    pause(0.5);

    [numRayPathsVal, statusNum] = lt.DbGet(forwardSimKey, 'RayNumberInVisiblePaths');
    if statusNum ~= DBGET_SUCCESS_CODE || isempty(numRayPathsVal) || ~isnumeric(numRayPathsVal)
        error('Could not get RayNumberInVisiblePaths. Ensure simulation has run.');
    end
    numRayPaths = double(numRayPathsVal);
    if numRayPaths == 0, error('No ray paths found (RayNumberInVisiblePaths is 0).'); end
    fprintf('  Total ray paths to process: %d.\n', numRayPaths);
    
    fprintf('  Retrieving data for all %d ray paths...\n', numRayPaths);
    allRayData = cell(numRayPaths, 4);
    for i = 1:numRayPaths
        allRayData{i,1} = i;
        [val, statP] = lt.DbGet(forwardSimKey, 'RayPathPowerAt', i, 1);
        if statP == DBGET_SUCCESS_CODE, allRayData{i,2} = double(val); else allRayData{i,2} = NaN; end
        [val, statS] = lt.DbGet(forwardSimKey, 'RayPathSourceNameAt', i, 1);
        if statS == DBGET_SUCCESS_CODE, allRayData{i,3} = char(val); else allRayData{i,3} = ''; end
        [val, statF] = lt.DbGet(forwardSimKey, 'RayPathFinalSurfaceAt', i, 1);
        if statF == DBGET_SUCCESS_CODE, allRayData{i,4} = char(val); else allRayData{i,4} = ''; end
    end
    fprintf('  Finished retrieving all ray path data.\n');
end

function [sourceFilt, surfaceFilt, cancelled] = getInteractiveFilters(allRayData)
    cancelled = false;
    validDataIdx = ~cellfun('isempty', allRayData(:,3)) & ~cellfun('isempty', allRayData(:,4));
    uniqueSourceNames = [{'* (All Sources)'}; unique(allRayData(validDataIdx,3))];
    uniqueFinalSurfaces = [{'* (All Surfaces)'}; unique(allRayData(validDataIdx,4))];
    
    [selIdx, ok] = listdlg('ListString', uniqueSourceNames, 'SelectionMode', 'single', 'Name', 'Select Source Filter', 'PromptString', 'Select a source name:');
    if ~ok, cancelled = true; disp('Source selection cancelled.'); return; end
    sourceFilt = uniqueSourceNames{selIdx}; if strcmp(sourceFilt, '* (All Sources)'), sourceFilt = '*'; end
    
    [selIdx, ok] = listdlg('ListString', uniqueFinalSurfaces, 'SelectionMode', 'single', 'Name', 'Select Final Surface Filter', 'PromptString', 'Select a final surface:');
    if ~ok, cancelled = true; disp('Final surface selection cancelled.'); return; end
    surfaceFilt = uniqueFinalSurfaces{selIdx}; if strcmp(surfaceFilt, '* (All Surfaces)'), surfaceFilt = '*'; end
end

function [sortedData, totalPower] = filterAndSortRayData(allRayData, sourceFilt, surfFilt)
    validIdx = ~cellfun(@(x) any(isnan(x)), allRayData(:,2));
    filteredData = allRayData(validIdx,:);
    if ~strcmp(sourceFilt, '*'), matchIdx = strcmp(filteredData(:,3), sourceFilt); filteredData = filteredData(matchIdx,:); end
    if ~strcmp(surfFilt, '*'), matchIdx = strcmp(filteredData(:,4), surfFilt); filteredData = filteredData(matchIdx,:); end
    if isempty(filteredData), error('No ray paths match the selected filters.'); end
    fprintf('  %d paths remaining after filtering.\n', size(filteredData,1));
    
    powers = cell2mat(filteredData(:,2));
    [~, sortOrder] = sort(powers, 'descend');
    sortedData = filteredData(sortOrder,:);
    totalPower = sum(powers(~isnan(powers)));
end

function topIndices = getRayIndicesForCumulativePercent(sortedRayData, totalPower, percentTarget)
    topIndices = [];
    if totalPower <= 1e-12 || percentTarget <= 1e-9, return; end
    if percentTarget >= 100, topIndices = cell2mat(sortedRayData(:,1)); return; end

    powerThresholdAbsolute = (percentTarget / 100) * totalPower;
    cumulativePower = 0;
    
    for i_ray = 1:size(sortedRayData, 1)
        rayPower = sortedRayData{i_ray, 2};
        if isnan(rayPower), continue; end
        cumulativePower = cumulativePower + rayPower;
        topIndices(end+1) = sortedRayData{i_ray, 1};
        if cumulativePower >= (powerThresholdAbsolute - 1e-9), break; end
    end
end

function setVisibility(lt, fwdSimKey, numPaths, indicesToShow)
    disp('    Setting ray path visibilities...');
    lt.Cmd('DBUpdateOff'); lt.Cmd('RecalcOff');
    for i = 1:numPaths, lt.DbSet(fwdSimKey, 'RayPathVisibleAt', "No", i, 1); end
    if ~isempty(indicesToShow)
        for i = 1:length(indicesToShow), lt.DbSet(fwdSimKey, 'RayPathVisibleAt', "Yes", indicesToShow(i), 1); end
    end
    lt.Cmd('DBUpdateOn'); lt.Cmd('RecalcOn');
    lt.Cmd("RecalcNow"); pause(2.5);
    lt.Cmd('ShowOnlyRayPathRays'); pause(1.0);
end

function saveIntervalData(allRayData, indices, saveDir, baseName, runID, interval, srcFilt, surfFilt)
    if isempty(indices), return; end
    
    intervalFileSuffix = sprintf('P%g-%g', interval(1), interval(2));
    matFileNameInterval = sprintf('%s_%s_%s_Src-%s_Surf-%s_Data.mat', ...
        baseName, runID, intervalFileSuffix, ...
        matlab.lang.makeValidName(srcFilt), matlab.lang.makeValidName(surfFilt));
    fullMatPathInterval = fullfile(saveDir, matFileNameInterval);
    
    topRaysDataToSave = cell(length(indices), size(allRayData,2));
    for i_save = 1:length(indices)
        originalIdx = indices(i_save);
        rowDataIdx = find(cell2mat(allRayData(:,1)) == originalIdx, 1);
        if ~isempty(rowDataIdx), topRaysDataToSave(i_save,:) = allRayData(rowDataIdx, :); end
    end
    
    if ~isempty(topRaysDataToSave)
        save(fullMatPathInterval, 'topRaysDataToSave', 'interval', 'sourceFilt', 'surfFilt');
        disp(['    Data for interval rays saved to: ' matFileNameInterval]);
    end
end

% =========================================================================
% Helper Function to Capture Active LightTools View via Clipboard
% =========================================================================
function success = captureViaClipboard(ltHandle, fullSavePath, imageFormat)
    % Captures the currently active view in LightTools by using the
    % 'CopyToClipboard' command and then saving the clipboard content as
    % an image file.

    success = false; % Default to failure
    fprintf('\n--- Capturing View via Clipboard to %s ---\n', fullSavePath);

    % Define success code for lt.Cmd - adjust if your API version is different
    DB_CMD_SUCCESS = 0; 
    
    try
        % --- 1. Ensure a view is active (the caller should have done this) ---
        % But we add a small pause to ensure UI is ready for the next command.
        pause(0.5);

        % --- 2. Execute CopyToClipboard ---
        fprintf('  Executing CopyToClipboard...\n');
        statusCopy = ltHandle.Cmd('CopyToClipboard');
        if statusCopy ~= DB_CMD_SUCCESS
            % Attempt to get a more descriptive error message
            try errStr = char(ltHandle.GetStatusString(statusCopy)); catch; errStr = 'Unknown'; end
            warning('  CopyToClipboard command may have failed. Status: %d (%s)', statusCopy, errStr);
            % We don't exit here because something might still be on the clipboard
        else
            disp('  CopyToClipboard command sent successfully.');
        end
        pause(1.0); % CRUCIAL: Pause to allow the OS clipboard to be updated.

        % --- 3. Retrieve Image from Clipboard and Save using Java AWT ---
        fprintf('  Attempting to save image from clipboard...\n');
        
        clipboard = java.awt.Toolkit.getDefaultToolkit().getSystemClipboard();
        transferable = clipboard.getContents(java.lang.Object); % Get clipboard contents

        % Check if the content is an image
        if ~isempty(transferable) && transferable.isDataFlavorSupported(java.awt.datatransfer.DataFlavor.imageFlavor)
            awtImage = transferable.getTransferData(java.awt.datatransfer.DataFlavor.imageFlavor);
            
            % Handle different potential Java image types returned by the clipboard
            if isa(awtImage, 'java.awt.image.BufferedImage')
                % The object is already a BufferedImage, which is ideal
                bufferedImage = awtImage;
            else
                % The object is a standard AWT Image, which needs to be drawn onto a
                % BufferedImage before it can be saved by ImageIO.
                
                % Get image dimensions using ImageObserver. Pass [] for null.
                imageObserver = java.awt.image.ImageObserver.empty;
                width = awtImage.getWidth(imageObserver);
                height = awtImage.getHeight(imageObserver);
                
                % Sometimes dimensions are not immediately available (-1). Wait and retry.
                retry_dim = 0; max_retry_dim = 5;
                while (width == -1 || height == -1) && retry_dim < max_retry_dim
                    pause(0.2);
                    width = awtImage.getWidth(imageObserver);
                    height = awtImage.getHeight(imageObserver);
                    retry_dim = retry_dim + 1;
                end
                
                if width == -1 || height == -1
                    error('Could not determine image dimensions from clipboard data after retries.');
                end
                
                % Create a new BufferedImage and draw the AWT image onto it
                bufferedImage = java.awt.image.BufferedImage(width, height, java.awt.image.BufferedImage.TYPE_INT_RGB);
                graphics = bufferedImage.createGraphics();
                graphics.drawImage(awtImage, 0, 0, imageObserver);
                graphics.dispose(); % Release graphics resources
            end

            % Save the BufferedImage to a file
            fileOutput = java.io.File(fullSavePath);
            successSave = javax.imageio.ImageIO.write(bufferedImage, upper(imageFormat), fileOutput);
            
            if successSave
                disp(['    Successfully saved: ' fullSavePath]);
                success = true;
            else
                warning('    Failed to save image from clipboard. javax.imageio.ImageIO.write returned false. Check image format and path.');
            end
        else
            warning('    No image data found on the clipboard, or data flavor not supported.');
            
            % For debugging: check what data flavors ARE available
            if ~isempty(transferable)
                availableFlavors = transferable.getTransferDataFlavors();
                disp('    Available clipboard data flavors:');
                for i_flav = 1:length(availableFlavors)
                    try
                        flavorStr = char(availableFlavors(i_flav).toString());
                        disp(['      ' num2str(i_flav) ': ' flavorStr]);
                    catch
                        disp(['      ' num2str(i_flav) ': <could not convert flavor to string>']);
                    end
                end
            else
                disp('    Clipboard transferable object is empty.');
            end
        end
        
    catch ME_clipboard
        warning('    Error during clipboard capture: %s', ME_clipboard.message);
        fprintf(2, '    Clipboard error in %s at line %d\n', ME_clipboard.stack(1).file, ME_clipboard.stack(1).line);
    end
    
    fprintf('--- Finished Clipboard Capture ---\n\n');
end

function restoreAllRayVisibility(lt, key, numPaths)
    if ~iscom(lt) || isempty(key) || numPaths == 0, return; end
    disp('--- Restoring visibility for all ray paths ---');
    lt.Cmd('DBUpdateOff'); lt.Cmd('RecalcOff');
    for i = 1:numPaths
        lt.DbSet(key, 'RayPathVisibleAt', "Yes", i, 1);
    end
    lt.Cmd('DBUpdateOn'); lt.Cmd('RecalcOn');
    lt.Cmd("RecalcNow");
    disp('All ray paths set to visible.');
end

function releaseLightTools(lt)
    if ~isempty(lt) && iscom(lt), lt.delete; end
    disp('LightTools connection released.');
end

% =========================================================================
% Helper Function to Parse a String of Intervals (e.g., [[100,70],[50,0]])
% =========================================================================
function [intervals, err] = parseIntervalString(str)
    % Parses a string like '[[100,70],[70,0]]' into a MATLAB cell array
    % where each cell contains a 1x2 vector, e.g., {[100 70], [70 0]}.
    % Returns an error message string if parsing fails.

    intervals = {}; % Initialize as empty cell array
    err = '';       % Initialize empty error string

    try
        % 1. Preliminary cleanup: remove whitespace and the outer-most brackets
        % This simplifies parsing for both single and multiple intervals.
        cleanedStr = strtrim(str);
        if startsWith(cleanedStr, '[') && endsWith(cleanedStr, ']')
            cleanedStr = cleanedStr(2:end-1);
        end

        % If the string is now empty, there's nothing to parse.
        if isempty(cleanedStr)
            return;
        end

        % 2. Use a regular expression to find all occurrences of [num,num]
        % This pattern looks for:
        %   \[          - a literal opening bracket
        %   \s*         - zero or more whitespace characters
        %   (-?\d+\.?\d*) - captures the first number (integer or decimal, optional negative)
        %   \s*,\s*     - a comma, surrounded by optional whitespace
        %   (-?\d+\.?\d*) - captures the second number
        %   \s*         - zero or more whitespace characters
        %   \]          - a literal closing bracket
        intervalMatches = regexp(cleanedStr, '\[\s*(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)\s*\]', 'tokens');

        if isempty(intervalMatches)
            err = 'Could not find any valid [number, number] patterns in the input string.';
            return;
        end

        % 3. Loop through the matches and build the cell array
        for k = 1:length(intervalMatches)
            currentMatch = intervalMatches{k}; % This is a cell {'num1_str', 'num2_str'}
            
            upperVal = str2double(currentMatch{1});
            lowerVal = str2double(currentMatch{2});
            
            % 4. Validate each parsed pair
            if isnan(upperVal) || isnan(lowerVal)
                err = sprintf('Interval #%d ("%s") contains non-numeric values.', k, strjoin(currentMatch,','));
                intervals = {}; % Clear partial results on error
                return;
            end
            
            if upperVal < lowerVal
                err = sprintf('Invalid interval [%g,%g]: Upper value must be >= lower value.', upperVal, lowerVal);
                intervals = {};
                return;
            end
            
            if upperVal > 100 || lowerVal < 0
                err = sprintf('Invalid interval [%g,%g]: Values must be between 0 and 100.', upperVal, lowerVal);
                intervals = {};
                return;
            end
            
            % If valid, add to the output cell array
            intervals{end+1} = [upperVal, lowerVal];
        end
        
    catch ME
        err = sprintf('An unexpected error occurred during parsing: %s', ME.message);
        intervals = {};
    end
end
