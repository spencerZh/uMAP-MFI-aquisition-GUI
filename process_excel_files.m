function process_excel_files()
    % Ask user for inputs
    maxColumn = str2double(input('Enter max column (e.g., 7): ', 's'));
    antigen = str2double(input('Enter antigen column number (e.g., 14): ', 's'));
    antigenLabels = {'OV16', 'OVFABP', 'OVMSA', 'OV53-74', 'BSA', 'BM14', 'WB123', ...
                     'OV33', 'OV15', 'OV34-47', 'BM33', 'BMR1', 'BMSerpin', 'AG'};

    % Get current folder
    scriptDir = fileparts(mfilename('fullpath'));
    files = dir(fullfile(scriptDir, '*.xlsx'));

    % Loop through each file
    for f = 1:length(files)
        filePath = fullfile(scriptDir, files(f).name);

        % Process 2nd and 3rd sheets
        for sheetIdx = 2:3
            try
                data = readcell(filePath, 'Sheet', sheetIdx);
            catch
                fprintf('Skipping sheet %d in file %s due to read error.\n', sheetIdx, files(f).name);
                continue;
            end

            currentRow = 4;
            numRows = size(data, 1);

            while (currentRow + 2) <= numRows && ~isempty(data{currentRow, 1})
                % Copy block only if full 3 rows available
                for r = 0:2
                    for c = 1:maxColumn
                        % Ensure column exists
                        if size(data, 2) < c
                            continue;
                        end
                        srcVal = data{currentRow + r, c};
                        destCol = maxColumn + c;
                        data{r + 1, destCol} = srcVal;
                    end
                end
                currentRow = currentRow + 3;
            end

            % Compute column-wise mean of rows 1–3 for columns 1:min(antigen, availableCols)
            availableCols = size(data, 2);
            for c = 1:min(antigen, availableCols)
                vals = data(1:3, c);
                try
                    numericVals = cellfun(@(x) double(x), vals);
                    data{10, c} = mean(numericVals, 'omitnan');
                catch
                    data{10, c} = NaN;
                end
            end
            
            data(9, 1:14) = antigenLabels;
            
            % Write updated sheet back
            try
                writecell(data, filePath, 'Sheet', sheetIdx);
            catch
                fprintf('Failed to write sheet %d in file %s\n', sheetIdx, files(f).name);
            end
        end
    end

    fprintf('Processing complete.\n');
end
