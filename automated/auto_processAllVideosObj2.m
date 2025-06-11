function processAllVideosObj2(parentDir)
    % Get all .mp4 files in subdirectories
    allFiles = dir(fullfile(parentDir, '**', '*.mp4'));
    numFiles = length(allFiles);

    fprintf('Found %d video files.\n', numFiles);
    for idx = 1:numFiles
        try
            processSingleVideo(allFiles(idx));
        catch ME
            fprintf('Error processing file %s: %s\n', fullfile(allFiles(idx).folder, allFiles(idx).name), ME.message);
        end
    end
end

function processSingleVideo(fileInfo)
    videoFile = fullfile(fileInfo.folder, fileInfo.name);
    fprintf('Processing video: %s\n', videoFile);

    % Initialize video reader
    videoReader = VideoReader(videoFile);
    inputFPS = videoReader.FrameRate;

    % Load tracking data
    trackingFile = fullfile(fileInfo.folder, 'Tracking.mat');
    if ~isfile(trackingFile)
        warning('Tracking.mat not found in %s. Skipping.', fileInfo.folder);
        return;
    end
    load(trackingFile, 'Tracking');

    % Extract pose coordinates
    NoseX = Tracking.Smooth.Nose(1, :);
    NoseY = Tracking.Smooth.Nose(2, :);
    HeadX = Tracking.Smooth.Head(1, :);
    HeadY = Tracking.Smooth.Head(2, :);
    TailbaseX = Tracking.Smooth.Tailbase(1, :);
    TailbaseY = Tracking.Smooth.Tailbase(2, :);

    % Load behavior data
    behaviorFile = fullfile(fileInfo.folder, 'Behavior.mat');
    if ~isfile(behaviorFile)
        warning('Behavior.mat not found in %s. Skipping.', fileInfo.folder);
        return;
    end
    load(behaviorFile, 'Behavior');
    explorationBouts = Behavior.Exploration.Bouts;

    % Prepare output video
    [~, folderName] = fileparts(fileInfo.folder);
    outputVideoFile = fullfile(fileInfo.folder, [folderName, '_processed.mp4']);
    outputVideo = VideoWriter(outputVideoFile, 'MPEG-4');
    outputVideo.FrameRate = inputFPS;
    open(outputVideo);

    % Process frames
    frameIdx = 1;
    frameInterval = 1;
    while hasFrame(videoReader)
        frame = readFrame(videoReader);
        actualFrameIdx = frameIdx;

        if actualFrameIdx <= length(NoseX)
            frame = overlayMarkers(frame, NoseX(actualFrameIdx), NoseY(actualFrameIdx), ...
                                   HeadX(actualFrameIdx), HeadY(actualFrameIdx), ...
                                   TailbaseX(actualFrameIdx), TailbaseY(actualFrameIdx));
        end

        if any(actualFrameIdx >= explorationBouts(:, 1) & actualFrameIdx <= explorationBouts(:, 2))
            frame = insertText(frame, [150, 0], 'Explore_Obj2', 'FontSize', 26, ...
                'BoxColor', 'blue', 'BoxOpacity', 1, 'TextColor', 'white');
        end

        writeVideo(outputVideo, frame);
        frameIdx = frameIdx + frameInterval;
    end

    close(outputVideo);
    fprintf('Finished processing: %s\n', outputVideoFile);
end

function frame = overlayMarkers(frame, nx, ny, hx, hy, tx, ty)
    frame = insertShape(frame, 'FilledCircle', [nx, ny, 2], 'Color', 'yellow', 'Opacity', 1);
    frame = insertShape(frame, 'FilledCircle', [hx, hy, 2], 'Color', 'red', 'Opacity', 1);
    frame = insertShape(frame, 'FilledCircle', [tx, ty, 2], 'Color', 'cyan', 'Opacity', 1);
end