clc;
clear;
close all;

% Define the directory containing the video files
videoDir = '/Replace/with/correct/video_directory_path';

if ~isfolder(videoDir)
    error('The specified video directory does not exist: %s', videoDir);
end

sceneFiles = dir(fullfile(videoDir, '*Obj2.mp4'));

for k = 1:length(sceneFiles)
    try
        sceneVideoFile = fullfile(sceneFiles(k).folder, sceneFiles(k).name);
        objectVideoFile = strrep(sceneVideoFile, 'Obj2', 'Obj1');
        
        if ~isfile(objectVideoFile)
            warning('Object video file not found for %s. Skipping.', sceneVideoFile);
            continue;
        end

        outputVideoFile = strrep(sceneVideoFile, 'Obj2', 'Overlay');
        sceneVid = VideoReader(sceneVideoFile);
        objectVid = VideoReader(objectVideoFile);
        inputFPS = sceneVid.FrameRate;

        outputVid = VideoWriter(outputVideoFile, 'MPEG-4');
        outputVid.FrameRate = inputFPS;
        open(outputVid);

        rowShift = 1; colShift = 1;

        while hasFrame(sceneVid) && hasFrame(objectVid)
            sceneFrame = readFrame(sceneVid);
            objectFrame = readFrame(objectVid);

            [sceneRows, sceneCols, ~] = size(sceneFrame);
            [objRows, objCols, ~] = size(objectFrame);
            endRow = min(sceneRows, rowShift + objRows - 1);
            endCol = min(sceneCols, colShift + objCols - 1);

            sceneFrame(rowShift:endRow, colShift:endCol, :) = ...
                objectFrame(1:(endRow - rowShift + 1), 1:(endCol - colShift + 1), :);

            writeVideo(outputVid, sceneFrame);
        end

        close(outputVid);
        fprintf('Processed and saved: %s\n', outputVideoFile);
        
    catch ME
        warning('Error processing %s: %s', sceneFiles(k).name, ME.message);
    end
end