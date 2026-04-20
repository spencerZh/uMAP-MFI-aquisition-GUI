function Combine_Median()
    % Combine_Median: Combines antigen data and BSA-subtracted values into a new Excel file
    % Includes:
    %   - Sheet1: No Subtract
    %   - Sheet2: BSA Subtract
    %   - Sheet3: dot1 replacement (BSA Subtract with negatives replaced by 0.1)

    % Global parameters
    global antigen_amount antigen_name_row_index MFI
    global BSA_column_index Perform_BSA_subtract
    antigen_amount = 14;
    antigen_name_row_index = 9;
    MFI = 10;
    BSA_column_index = 5;
    Perform_BSA_subtract = 1;

    % Get directory of this function
    funcDir = fileparts(mfilename('fullpath'));

    % Get all Excel files
    allFiles = dir(fullfile(funcDir, '*.xlsx'));
    fileNames = {allFiles.name};
    validMask = ~startsWith(fileNames, '~$');
    excelFiles = allFiles(validMask);

    if isempty(excelFiles)
        disp('No valid Excel files found.');
        return;
    end

    % Sort files by numeric "cX" in filename
    colNums = zeros(1, numel(excelFiles));
    pattern = '_c(\d+)_ROI';
    for i = 1:numel(excelFiles)
        tokens = regexp(excelFiles(i).name, pattern, 'tokens');
        if ~isempty(tokens)
            colNums(i) = str2double(tokens{1}{1});
        else
            colNums(i) = Inf;
        end
    end
    [~, sortIdx] = sort(colNums);
    excelFiles = excelFiles(sortIdx);

    % Read header (antigen names) from first file
    firstFile = fullfile(funcDir, excelFiles(1).name);
    [~, ~, rawAntigens] = xlsread(firstFile, 3);

    % Construct header row
    headerRow = cell(1, antigen_amount + 1);
    headerRow{1} = 'sample';
    for col = 1:antigen_amount
        headerRow{col + 1} = rawAntigens{antigen_name_row_index, col};
    end

    % Build sheet1 data (sample names + MFI values)
    numFiles = numel(excelFiles);
    dataRows = cell(numFiles, antigen_amount + 1);
    for i = 1:numFiles
        fileName = excelFiles(i).name;
        fullPath = fullfile(funcDir, fileName);
        rawData = readcell(fullPath, 'Sheet', 3);

        dataRows{i, 1} = fileName;  % Sample name
        for col = 1:antigen_amount
            dataRows{i, col + 1} = rawData{MFI, col};
        end
    end

    % Sheet 1: No Subtract
    sheet1 = [headerRow; dataRows];

    % Convert numeric part of dataRows to a matrix for vectorized processing
    numericData = cell2mat(dataRows(:, 2:end));  % numFiles x antigen_amount

    if Perform_BSA_subtract == 1
        % --- Sheet 2: BSA Subtract ---
        bsaValues = numericData(:, BSA_column_index);        % BSA column vector
        numericBsaSubtracted = numericData - bsaValues;     % vectorized subtraction
        sheet2 = [dataRows(:,1), num2cell(numericBsaSubtracted)];
        sheet2 = [headerRow; sheet2];

        % --- Sheet 3: dot1 replacement ---
        numericDot1 = numericBsaSubtracted;
        numericDot1(numericDot1 < 0) = 0.1;                 % vectorized replacement
        dot1Data = [dataRows(:,1), num2cell(numericDot1)];
        dot1Data = [headerRow; dot1Data];
    else
        sheet2 = {'No BSA Subtraction Performed'};
        dot1Data = {'No BSA Subtraction Performed'};
    end

    % Write all sheets to Excel
    outputFile = fullfile(funcDir, 'Combined_Median.xlsx');
    writecell(sheet1, outputFile, 'Sheet', 'No Subtract');
    writecell(sheet2, outputFile, 'Sheet', 'BSA Subtract');
    writecell(dot1Data, outputFile, 'Sheet', 'dot1 replacement');

    disp(['Combined_Median.xlsx saved with sheets: "No Subtract", "BSA Subtract", and "dot1 replacement".']);
end
