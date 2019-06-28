clear all
clc
close all

%% Set the folder's name
dirs=uigetdir(dir(''));
dircell=struct2cell(dirs);
dirnum=length(dirs)-2;
C = string(dircell(2,1));
C = strsplit(C, '\');
CL = length(C);
%% Set the Excel file name
filename = C(CL-1)+' '+C(CL)+' Single'+'.xlsx'
%% Write in Excel
xlswrite(filename,{'Stress Level'},1,'A1')
xlswrite(filename,{'Time Point'},1,'B1')
xlswrite(filename,{'Cell Component'},1,'C1')

xlswrite(filename,{'Image ID'},1,'D1')
xlswrite(filename,{'x1'},1,'E1')
xlswrite(filename,{'y1'},1,'F1')
xlswrite(filename,{'x2'},1,'G1')
xlswrite(filename,{'y2'},1,'H1')
xlswrite(filename,{'Entropy'},1,'I1')
xlswrite(filename,{'Correlation Coefficient'},1,'J1')
xlswrite(filename,{'Energy'},1,'K1')

xlswrite(filename,{'Area'},1,'L1')
xlswrite(filename,{'Centroid'},1,'M1')
xlswrite(filename,{'BoundingBox'},1,'N1')
xlswrite(filename,{'MajorAxisLength'},1,'O1')
xlswrite(filename,{'MinorAxisLenth'},1,'P1')

xlswrite(filename,{'Eccentricty'},1,'Q1')
xlswrite(filename,{'Orientation'},1,'R1')
xlswrite(filename,{'ConvexArea'},1,'S1')
xlswrite(filename,{'Ratio of Image by ConvexArea'},1,'T1')

xlswrite(filename,{'FilledArea'},1,'U1')
xlswrite(filename,{'EulerNumber'},1,'V1')
xlswrite(filename,{'Extrema'},1,'W1')

xlswrite(filename,{'EquivDiameter'},1,'X1')
xlswrite(filename,{'Solidity'},1,'Y1')
xlswrite(filename,{'Extent'},1,'Z1')
xlswrite(filename,{'Perimeter'},1,'AA1')
xlswrite(filename,{'PerimeterOld'},1,'AB1')
%% Read each image

for i =1:dirnum
    A=strcat(dircell(2,2+i),'\',dircell(1,2+i));
    Name = string(A)
    A1 = strcat('A',num2str(i+1));
    xlswrite(filename,{num2str(C(CL-2))},1,A1)
    B1 = strcat('B',num2str(i+1));
    xlswrite(filename,{num2str(C(CL-1))},1,B1)
    C1 = strcat('C',num2str(i+1));
    xlswrite(filename,{num2str(C(CL))},1,C1)
    %% Analysis
    D1 = strcat('D',num2str(i+1));
    xlswrite(filename,dircell(1,2+i),1,D1)
    %% read the image
    Image = imread(Name);
    figure(1), imshow(Image);title('Original Image');
    %% Adaptive filtering
    I_eq = adapthisteq(Image);
    figure(2), imshow(I_eq);title('After adaptive histogram');
    %% Binarize the Image
    bw = imbinarize(I_eq);
    bw = bwareaopen(bw,200);
    figure(3), imshow(bw);title('Binary image');
    %% Have user specify the area they want to define as neutral colored (white  or gray).
    promptMessage = sprintf('Drag out a box over Single Cell  ROI you want to be .\nDouble-click inside of it to finish it.');
    titleBarCaption = 'Continue?';
    button = questdlg(promptMessage, titleBarCaption, 'Draw', 'Cancel', 'Draw');
    if strcmpi(button, 'Cancel')
        return;
    end
    %% CROPPING WITH DEFINING THE BOX
    hBox = imrect;
    roiPosition = wait(hBox);   % Wait for user to double-click
    roiPosition  % Display in command window.
    E1 = strcat('E',num2str(i+1));
    xlswrite(filename,roiPosition(1),1,E1)
    F1 = strcat('F',num2str(i+1));
    xlswrite(filename,roiPosition(2),1,F1)
    G1 = strcat('G',num2str(i+1));
    xlswrite(filename,roiPosition(3),1,G1)
    H1 = strcat('H',num2str(i+1));
    xlswrite(filename,roiPosition(4),1,H1)
    %% Get box coordinates so we can crop a portion out of the full sized image.
    xCoords = [roiPosition(1), roiPosition(1)+roiPosition(3), roiPosition(1)+roiPosition(3), roiPosition(1), roiPosition(1)];
    yCoords = [roiPosition(2), roiPosition(2), roiPosition(2)+roiPosition(4), roiPosition(2)+roiPosition(4), roiPosition(2)];
    croppingRectangle = roiPosition;
    %% CROP PORTION
    cropPortion = imcrop(bw, croppingRectangle);
    figure, imshow(cropPortion), title('crop region');
    I=cropPortion;
    figure(4);imshow(I);title('Cropped Image');
    %% filling the holes
    bw2 = imfill(I, 'holes');
    bw3 = imopen (bw2, ones(5,5));
    bw4 = bwareaopen(bw3, 40);
    bw4_perim = bwperim(bw4);
    overlay1 = imoverlay(I, bw4_perim, [.3 1 .3]);
    figure(5), imshow(overlay1);title('Overlayed Image');
    %% Edge detection
    [~, threshold] = edge(bw4_perim, 'sobel');
    fudgeFactor = 2;
    BWs = edge(bw4_perim,'sobel', threshold * fudgeFactor);
    figure (6), imshow(BWs), title('binary gradient mask');
    %% Close the borders
    bw2 = imfill(BWs, 'holes');
    bw3 = imclose (bw2, ones(5,5));
    bw4 = bwareaopen(bw3, 100);
    bw4_perim = bwperim(bw4);
    overlay1 = imoverlay(BWs, bw4_perim, [.3 1 .3]);
    figure(7), imshow(bw4_perim);
    figure(8), imshow(overlay1);
    mask_em = imextendedmax(I_eq, 20);
    %% Masking
    mask_em = imclose(I, ones(5,5));
    mask_em = imclose(mask_em, ones(5,5));
    mask_em=imfill(mask_em,'holes');
    mask_em = bwareaopen(mask_em,1000);
    figure(9), imshow(mask_em);

    %% rotating the images to saame orientationn 
    measurements = regionprops(bw4_perim, 'Orientation');
    angle = measurements(1).Orientation;
    angleToRotateBy = 90 - angle; % or "+" or 180 (instead of 90), you'd have to check.
    rotatedImage = imrotate(bw4_perim, angleToRotateBy);
    verticalProfile = sum(rotatedImage, 2);
    firstLine = find(verticalProfile > 0, 1, 'first');
    lastLine = find(verticalProfile > 0, 1, 'last');
    topSum = sum(verticalProfile(firstLine:firstLine+50));
    bottomSum = sum(verticalProfile(lastLine-50:lastLine));
    if topSum > bottomSum
      rotatedImage = imrotate(rotatedImage, 180);
    end
    figure(10),imshow(rotatedImage);title('rotating to the same orientation');
    %% CROPPING from the original image
    crop = imcrop(Image, roiPosition);
    figure(11), imshow(crop);
    %% Multiplying it to the mask
    G = uint16(crop) .*uint16(mask_em);
    figure(12), imshow(G)
    %% region props calculation
    L = bwlabel(G);
    stats = regionprops(L,'all');%%%
    L1 = strcat('L',num2str(i+1));
    xlswrite(filename,stats.Area,1,L1)
    
    M1 = strcat('M',num2str(i+1));
    xlswrite(filename,{strcat(num2str(stats.Centroid(1,1)),',',num2str(stats.Centroid(1,2)))},1,M1)
    
    N1 = strcat('N',num2str(i+1));
    xlswrite(filename,{strcat(num2str(stats.BoundingBox(1,1)),',',num2str(stats.BoundingBox(1,2)),',',num2str(stats.BoundingBox(1,3)),',',num2str(stats.BoundingBox(1,4)))},1,N1)
    
    O1 = strcat('O',num2str(i+1));
    xlswrite(filename,stats.MajorAxisLength,1,O1)
    
    P1 = strcat('P',num2str(i+1));
    xlswrite(filename,stats.MinorAxisLength,1,P1)
    
    Q1 = strcat('Q',num2str(i+1));
    xlswrite(filename,stats.Eccentricity,1,Q1)
    
    R1 = strcat('R',num2str(i+1));
    xlswrite(filename,stats.Orientation,1,R1)
    
    S1 = strcat('S',num2str(i+1));
    xlswrite(filename,stats.ConvexArea,1,S1)
    
    T1 = strcat('T',num2str(i+1));
    [LEN, HEI]=size(Image);
    AREA_OF_IMAGE = LEN*HEI;
    xlswrite(filename,AREA_OF_IMAGE/stats.ConvexArea,1,T1)
    
    U1 = strcat('U',num2str(i+1));
    xlswrite(filename,stats.FilledArea,1,U1)
    
    V1 = strcat('V',num2str(i+1));
    xlswrite(filename,stats.EulerNumber,1,V1)
    
    %W1 = strcat('W',num2str(i+1));
    %xlswrite(filename,stats.Extrema,1,W1)
    
    X1 = strcat('X',num2str(i+1));
    xlswrite(filename,stats.EquivDiameter,1,X1)
    
    Y1 = strcat('Y',num2str(i+1));
    xlswrite(filename,stats.Solidity,1,Y1)
    
    Z1 = strcat('Z',num2str(i+1));
    xlswrite(filename,stats.Extent,1,Z1)
    
    AA1 = strcat('AA',num2str(i+1));
    xlswrite(filename,stats.Perimeter,1,AA1)
    
    AB1 = strcat('AB',num2str(i+1));
    xlswrite(filename,stats.PerimeterOld,1,AB1)
    %% entropy
    entropy_of_single_cell= entropy(G)
    I1 = strcat('I',num2str(i+1));
    xlswrite(filename,entropy_of_single_cell,1,I1)
    %% correlation of the cell
    Correlation=stdfilt(G);
    correlation_of_single_cell=corr2(G,Correlation)
    J1 = strcat('J',num2str(i+1));
    xlswrite(filename,correlation_of_single_cell,1,J1)
    %% Energy of the cell
    energy_of_single_cell= graycoprops(G, {'energy'})
    K1 = strcat('K',num2str(i+1));
    xlswrite(filename,energy_of_single_cell.Energy,1,K1)
    %% Done
    pause(0.1)
    clc;
    close all;
end
