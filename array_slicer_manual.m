function array_slicer_manual()
    % Open a file dialog to select a TIFF image
    [file, path] = uigetfile({'*.tif;*.tiff','TIFF Files'}, 'Select a 16-bit TIFF image');
    if isequal(file, 0)
        disp('User canceled file selection.');
        return;
    end

    % Read the image
    img_path = fullfile(path, file);
    img = imread(img_path);

    % Display image size
    [img_height, img_width] = size(img);
    fprintf('Image size: %d rows x %d columns\n', img_height, img_width);

    % Prompt user for cropping parameters
    prompt = {'Enter origin_x (column index):', ...
              'Enter origin_y (row index):', ...
              'Enter array_column:', ...
              'Enter array_row:'};
    dlgtitle = 'Input Cropping Parameters';
    dims = [1 40];
    definput = {'770', '750', '14', '4'};
    answer = inputdlg(prompt, dlgtitle, dims, definput);

    if isempty(answer)
        disp('User canceled parameter input.');
        return;
    end

    % Convert inputs to numeric
    origin_x = str2double(answer{1});
    origin_y = str2double(answer{2});
    array_column = str2double(answer{3});
    array_row = str2double(answer{4});

    fprintf('origin_x = %d, origin_y = %d\n', origin_x, origin_y);
    fprintf('array_column = %d, array_row = %d\n', array_column, array_row);

    % Create root folder name from image file (without extension)
    [~, image_name, ~] = fileparts(file);
    root_folder = fullfile(path, image_name);

    % Run cropping routine
    crop_image_grid_manual(img, origin_x, origin_y, array_column, array_row, root_folder);

    % Open the root folder automatically
    if ispc
        winopen(root_folder);
    elseif ismac
        system(['open ', root_folder]);
    elseif isunix
        system(['xdg-open ', root_folder]);
    end
end

function crop_image_grid_manual(img, origin_x, origin_y, array_column, array_row, root_folder)
    crop_size = 550;           % 600 x 600 crop size
    crop_offset = 100;         % Offset from origin
    spacing = 900;             % Spacing between crops

    crop_start_x = origin_x - crop_offset;
    crop_start_y = origin_y - crop_offset;

    [img_height, img_width] = size(img);

    % Validate start point
    if crop_start_x < 1 || crop_start_y < 1
        warning('Out of bound: crop_start_x or crop_start_y < 1');
        return;
    end

    % Create root folder if it doesn't exist
    if ~exist(root_folder, 'dir')
        mkdir(root_folder);
    end

    row_count = 1;

    while row_count <= array_row
        fprintf('row: %d\n', row_count);
        column_count = 1;
        current_y = crop_start_y + spacing * (row_count - 1);

        if current_y + crop_size - 1 > img_height
            warning('Out of bound: crop region exceeds image height at row %d', row_count);
            return;
        end

        folder_name = fullfile(root_folder, sprintf('row_%d', row_count));
        if ~exist(folder_name, 'dir')
            mkdir(folder_name);
        end

        while column_count <= array_column
            current_x = crop_start_x + spacing * (column_count - 1);

            if current_x + crop_size - 1 > img_width
                warning('Out of bound: crop region exceeds image width at column %d', column_count);
                return;
            end

            % Crop and save
            crop_img = img(current_y : current_y + crop_size - 1, ...
                           current_x : current_x + crop_size - 1);
            filename = fullfile(folder_name, ...
                sprintf('crop_r%d_c%d.tif', row_count, column_count));
            imwrite(crop_img, filename, 'tif');
            fprintf('Saved: %s\n', filename);

            column_count = column_count + 1;
        end

        row_count = row_count + 1;
    end
end

