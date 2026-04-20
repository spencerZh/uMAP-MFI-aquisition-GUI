function roi_grid_measurement_tool()
    % ROI Grid Measurement Tool with Excel export and live control (R2023 compatible)

    % Defaults
    startX = 100; startY = 100; radius = 11;
    spacing = 60; nRows = 6; nCols = 7;
    previousMoveX = 0; previousMoveY = 0;

    % State
    roiArray = gobjects(nRows, nCols);
    img = [];
    imgHandle = [];
    imageFileName = '';
    imagePath = '';

    % Create GUI
    screenSize = get(groot, 'ScreenSize');
    guiWidth = min(1400, screenSize(3) - 100);
    guiHeight = min(800, screenSize(4) - 100);
    figLeft = (screenSize(3) - guiWidth)/2;
    figBottom = (screenSize(4) - guiHeight)/2;
    fig = uifigure('Name','ROI Grid Tool', 'Position',[figLeft figBottom guiWidth guiHeight]);

    ax = uiaxes(fig,'Position',[300 100 1050 700]);
    title(ax, 'Select an Image to Start');

    % Load/Save buttons
    btnLoad = uibutton(fig, 'Text','Load Image', 'Position',[20, 720, 120, 30], 'ButtonPushedFcn', @onLoadImage);
    btnCompute = uibutton(fig, 'Text','Compute Stats', 'Position',[20, 680, 120, 30], 'Enable','off', 'ButtonPushedFcn', @onComputeStats);
    btnSaveROI = uibutton(fig, 'Text','Save ROI', 'Position',[20, 640, 120, 30], 'Enable','off', 'ButtonPushedFcn', @onSaveROI);
    btnLoadROI = uibutton(fig, 'Text','Load ROI', 'Position',[20, 600, 120, 30], 'Enable','off', 'ButtonPushedFcn', @onLoadROI);
    btnExport = uibutton(fig, 'Text','Export to Excel', 'Position',[20, 560, 120, 30], 'Enable','off', 'ButtonPushedFcn', @onExportExcel);

    % Contrast
    uilabel(fig, 'Text', 'Contrast', 'Position', [20 510 60 20]);
    contrastSlider = uislider(fig, 'Limits', [1 65535], 'Value', 30000, ...
        'Position', [20, 490, 200, 3], 'ValueChangedFcn', @onContrastChange);

    % ROI controls
    inputStartX = uispinner(fig, 'Position', [80 450 120 22], 'Limits', [0 2000], 'Value', startX, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Start X', 'Position', [20 450 60 20]);
    inputStartY = uispinner(fig, 'Position', [80 420 120 22], 'Limits', [0 2000], 'Value', startY, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Start Y', 'Position', [20 420 60 20]);
    inputRadius = uispinner(fig, 'Position', [80 390 120 22], 'Limits', [1 200], 'Value', radius, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Radius', 'Position', [20 390 60 20]);
    inputSpacing = uispinner(fig, 'Position', [80 360 120 22], 'Limits', [1 200], 'Value', spacing, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Spacing', 'Position', [20 360 60 20]);
    inputRows = uispinner(fig, 'Position', [80 330 120 22], 'Limits', [1 20], 'Value', nRows, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Rows', 'Position', [20 330 60 20]);
    inputCols = uispinner(fig, 'Position', [80 300 120 22], 'Limits', [1 20], 'Value', nCols, 'ValueChangedFcn', @(~,~)resetROIs());
    uilabel(fig, 'Text', 'Cols', 'Position', [20 300 60 20]);

    % Move offsets
    inputMoveX = uispinner(fig, 'Position', [80 240 120 22], 'Limits', [-1000 1000], 'Value', 0, 'ValueChangedFcn', @(~,~)moveAllROIs());
    uilabel(fig, 'Text', 'Move X', 'Position', [20 240 60 20]);
    inputMoveY = uispinner(fig, 'Position', [80 210 120 22], 'Limits', [-1000 1000], 'Value', 0, 'ValueChangedFcn', @(~,~)moveAllROIs());
    uilabel(fig, 'Text', 'Move Y', 'Position', [20 210 60 20]);

    %% Callbacks

    function onLoadImage(~,~)
        [file, path] = uigetfile({'*.tif;*.tiff','TIFF Files'}, 'Select TIFF Image');
        if isequal(file,0), return; end
        imageFileName = file;
        imagePath = path;
        img = imread(fullfile(path, file));
    
        % Display the image
        if isempty(imgHandle) || ~isvalid(imgHandle)
            imgHandle = imshow(img, 'Parent', ax);
        else
            imgHandle.CData = img;
        end
        hold(ax,'on')
        ax.CLim = [0 contrastSlider.Value];
        axis(ax, 'image');
        ax.Position = [300 100 1050 700];
        ax.XLimMode = 'auto';
        ax.YLimMode = 'auto';
        ax.DataAspectRatio = [1 1 1];
        ax.XTick = [];
        ax.YTick = [];
        ax.Box = 'on';
        title(ax, file);
    
        % Enable controls
        btnCompute.Enable = 'on';
        btnSaveROI.Enable = 'on';
        btnLoadROI.Enable = 'on';
        btnExport.Enable = 'on';
    
        % Only reset ROIs if none exist
        if isempty(roiArray) || ~all(ishandle(roiArray(:)))
            resetROIs();  % create fresh ROIs only if none exist
        end
        bringGUIToFront(fig);
    end


    function onContrastChange(src, ~)
        if ~isempty(img)
            ax.CLim = [0 src.Value];
        end
    end

    function resetROIs()
        if isempty(img), return; end
        startX = inputStartX.Value;
        startY = inputStartY.Value;
        radius = inputRadius.Value;
        spacing = inputSpacing.Value;
        nRows = inputRows.Value;
        nCols = inputCols.Value;

        delete(findall(ax, 'Type', 'images.roi.Circle'));
        roiArray = gobjects(nRows, nCols);
        for r = 1:nRows
            for c = 1:nCols
                cx = startX + (c-1)*spacing;
                cy = startY + (r-1)*spacing;
                roiArray(r,c) = drawcircle(ax, 'Center', [cx cy], 'Radius', radius, 'Color', 'r');
            end
        end

        inputMoveX.Value = 0;
        inputMoveY.Value = 0;
        previousMoveX = 0;
        previousMoveY = 0;
    end

    function moveAllROIs()
        moveX = inputMoveX.Value;
        moveY = inputMoveY.Value;
        dx = moveX - previousMoveX;
        dy = moveY - previousMoveY;
        for k = 1:numel(roiArray)
            roiArray(k).Center = roiArray(k).Center + [dx dy];
        end
        previousMoveX = moveX;
        previousMoveY = moveY;
    end

    function onComputeStats(~,~)
        nRows = inputRows.Value;
        nCols = inputCols.Value;
        for r = 1:nRows
            for c = 1:nCols
                roi = roiArray(r,c);
                mask = createMask(roi);
                vals = double(img(mask));
                fprintf('ROI(%d,%d): Mean = %.2f, Median = %.2f\n', r, c, mean(vals), median(vals));
            end
        end
    end

    function onSaveROI(~,~)
        roiData = struct();
        nRows = inputRows.Value;
        nCols = inputCols.Value;
        roiData.nRows = nRows;
        roiData.nCols = nCols;
        roiData.centers = zeros(nRows*nCols, 2);
        roiData.radii = zeros(nRows*nCols, 1);
        for k = 1:numel(roiArray)
            roiData.centers(k,:) = roiArray(k).Center;
            roiData.radii(k) = roiArray(k).Radius;
        end
        [file, path] = uiputfile('roiData.mat','Save ROI Data');
        if isequal(file,0), return; end
        save(fullfile(path, file), 'roiData');
        bringGUIToFront(fig);
    end

    function onLoadROI(~,~)
        [file, path] = uigetfile('*.mat','Select ROI File');
        if isequal(file,0), return; end
        s = load(fullfile(path, file));
        roiData = s.roiData;
        inputRows.Value = roiData.nRows;
        inputCols.Value = roiData.nCols;
        inputRadius.Value = roiData.radii(1);
        nRows = roiData.nRows;
        nCols = roiData.nCols;
        delete(findall(ax, 'Type', 'images.roi.Circle'));
        roiArray = gobjects(nRows,nCols);
        for k = 1:nRows*nCols
            roiArray(k) = drawcircle(ax, ...
                'Center', roiData.centers(k,:), ...
                'Radius', roiData.radii(k), ...
                'Color', 'r');
        end
        inputMoveX.Value = 0;
        inputMoveY.Value = 0;
        previousMoveX = 0;
        previousMoveY = 0;
        bringGUIToFront(fig);
    end

    function onExportExcel(~,~)
        if isempty(img), return; end
        nRows = inputRows.Value;
        nCols = inputCols.Value;
        flatData = [];
        meanArray = zeros(nRows,nCols);
        medianArray = zeros(nRows,nCols);
        for r = 1:nRows
            for c = 1:nCols
                roi = roiArray(r,c);
                mask = createMask(roi);
                vals = double(img(mask));
                m = mean(vals);
                md = median(vals);
                flatData = [flatData; r, c, m, md];
                meanArray(r,c) = m;
                medianArray(r,c) = md;
            end
        end
        [~, baseName, ~] = fileparts(imageFileName);
        defaultExcelName = [baseName, '_ROI_Measurements.xlsx'];
        filepath = fullfile(imagePath, defaultExcelName);
        headers = {'Row', 'Col', 'Mean', 'Median'};
        writecell([headers; num2cell(flatData)], filepath, 'Sheet','FlatData');
        writematrix(meanArray, filepath, 'Sheet','MeanArray');
        writematrix(medianArray, filepath, 'Sheet','MedianArray');
        bringGUIToFront(fig);
    end

    function bringGUIToFront(fh)
        % R2023-compatible workaround
        fh.Visible = 'off';
        drawnow;
        fh.Visible = 'on';
        drawnow;
    end
end
