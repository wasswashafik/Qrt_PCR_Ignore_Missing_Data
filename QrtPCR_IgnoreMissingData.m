function QrtPCR_IgnoreMissingData_2020Jul7()

%--------------------------------------------------------------------------
%This function interprets results of Quantitative Real-Time Polymerase Chain Reaction results for SARS-CoV-2 samples

%Inputs

%None - Although the user is asked upon execution to select the desired
%folder to operate in/analyze results from. The users is also prompted to
%enter the Ct, Negative control column, and Positive control Column

%Ct - Cycle threshold. A channel is determined to be positive if it passes
%the fluorescence threshold before this number of run cycles have passed

%Negative control column - The column label (1-12) on the 96 well plate corresponding to the negative control

%Positive control column - The column label (1-12) on the 96 well plate corresponding to the positive control

%Outputs

%Excel spreadsheet (filename_Processed.xlsx) containing a visual readout of
%test results. A grid representing a 96 well plate is generated with labels
%of "Pass" or "Inconclusive" for control wells, and labels of "+", "-", or
%"Inconclusive" for sample wells. These labels are highlighted in color for aesthetic
%purposes with green corresponding to "Pass" or "-" results, red corresponding to
%"+" results, and yellow corresponding to "Inconclusive" results
%--------------------------------------------------------------------------

%Select excel spreadsheets from the desired folder and navigate to this
%folder
directoryName = uigetdir; 
cd(directoryName)
files = dir('*.xls');

%Initialize well plate layout for display
plate = cell(10,14);
plate(2,3:end) = [{'1'},{'2'},{'3'},{'4'},{'5'},{'6'},{'7'},{'8'},{'9'},{'10'},{'11'},{'12'}];
plate(3:end,2) = [{'A'},{'B'},{'C'},{'D'},{'E'},{'F'},{'G'},{'H'}];
columnLookup =  'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
nWell = 96;
nRowPlate = 8;
nColPlate = 12;
CtCol = 9;

%Prompt the user to enter the cycle threshold, positive and negative control locations
prompt = {'Enter the Cycle Threshold:','Enter the Negative Control Column:', 'Enter the Positive Control Column:'};
dlg_title = 'Input';
num_lines = 1;
defaultans = {'36','1','12'}; %Default Answers
answer = inputdlg(prompt,dlg_title,num_lines,defaultans);
Ct = str2double(answer{1}); %Convert answers to numeric values for indexing
negControlCol = str2double(answer{2});
posControlCol = str2double(answer{3});

%Loop through all spreadsheets in the folder for processing
for file = files'
   
    fileName = file.name;
    
    %Find the "." in the file name. This will be used to append the
    %"_Processed taeg to the output file. Period location is necessary to
    %account for filenames with periods in them and excel preeadsheets with
    %both .xls and .xlsx file extensions
    dotLoc = find(fileName == '.');
    
    if length(dotLoc) > 1
        
        dotLoc = dotLoc(end);
        
    end
    
    %Initialize plate layout for each file
    newPlate = plate;
    
    %Perform data extraction from each file. Data from both "Results" and
    %"Amplification Data" sheets are required 
    [num,~,raw] = xlsread(fileName,'Results');
    [numAmp,~,rawAmp] = xlsread(fileName,'Amplification Data');
    
    %Find where results begin in the spreadsheet. Instrument provides
    %leading administrative information before displaying data.
    begin = 1;
    
    while isnan(raw{begin,9})
        
        begin = begin+1;
   
    end 
    
        
    %Initialize the decision matrix for diagnosing each well, including
    %well locations for each decision type based on the colors described above. 
    goodCells = cell(1,nWell);
    badCells = cell(1,nWell);
    inconclusiveCells = cell(1,nWell);
    goodCounter = 1;
    badCounter = 1;
    inconclusiveCounter = 1;
    decisionMat = cell(nRowPlate,nColPlate);
    posControlDecision = 2*ones(nRowPlate,2);
    negControlDecision = 2*ones(nRowPlate,2);

    % Make control decisions
    posControlWells = posControlCol:nColPlate:nWell;
    negControlWells = negControlCol:nColPlate:nWell;
    rowNdx = num(:,1);
    
    %Positive control
    count = 1;
    for i = posControlWells(posControlWells <= max(rowNdx))
        
        if ~isnan(find(i == rowNdx))
            
            %Place the control labels in the output plate
            currRowNdxPos = find(rowNdx == i);

            if i == min(posControlWells)

               newPlate(1,2+posControlCol) = raw(begin + currRowNdxPos(1),4);

            end

            if ~isnan(num(currRowNdxPos(1),CtCol)) && num(currRowNdxPos(1),CtCol) < Ct%FAM channel decision

                posControlDecision(count,1) = 1;

            elseif isnan(raw{begin+currRowNdxPos(1),CtCol})

                posControlDecision(count,1) = 2;
                
            else
                
                posControlDecision(count,1) = 0;

            end

            if ~isnan(num(currRowNdxPos(2),CtCol)) && num(currRowNdxPos(2),CtCol) < Ct%HEX channel decision

                posControlDecision(count,2) = 1;

            elseif isnan(raw{begin+currRowNdxPos(2),CtCol})

                posControlDecision(count,2) = 2;
                
            else
                
                posControlDecision(count,2) = 0;

            end

            count = count +1;
            
        end
        
    end
    
%     deletePos = posControlWells(posControlWells > max(rowNdx));
%     posControlDecision(end-length(deletePos)+1:end,1) = 2;
    
    %Negative control
    count = 1;
    for i = negControlWells(negControlWells <= max(rowNdx))
        
        if ~isnan(find(i == rowNdx))
        
            currRowNdxNeg = find(rowNdx == i);

            if i == min(negControlWells)

               newPlate(1,2+negControlCol) = raw(begin + currRowNdxNeg(1),4);

            end

            if ~isnan(num(currRowNdxNeg(1),CtCol)) && num(currRowNdxNeg(1),CtCol) < Ct%FAM channel decision

                negControlDecision(count,1) = 1;

            elseif isnan(raw{begin+currRowNdxNeg(1),CtCol})

                negControlDecision(count,1) = 2;
                
            else
                
                negControlDecision(count,1) = 0;

            end

            if ~isnan(num(currRowNdxNeg(2),CtCol)) && num(currRowNdxNeg(2),CtCol) < Ct%HEX channel decision

                negControlDecision(count,2) = 1;

            elseif isnan(raw{begin+currRowNdxNeg(2),CtCol})

                negControlDecision(count,2) = 2;
                
            else
                
                negControlDecision(count,2) = 0;

            end

            count = count + 1;
            
        end
        
    end
    
%     deleteNegs = negControlWells(negControlWells > max(rowNdx));
%     negControlDecision(end-length(deleteNegs)+1:end,1) = 2;
    
    % Create labels based on control decisions
    for i = 1:nRowPlate
        
        %Negative control - Any amplification seen crossing the fluorescence
        %threshold before CT triggers an "Inconclusive" result. No
        %amplification, or amplification after Ct = "Pass"
        if negControlDecision(i,1) == 1 || negControlDecision(i,2) == 1
            
            decisionMat(i,negControlCol) = {'Inconclusive'};
            
            inconclusiveCells(inconclusiveCounter) = {[columnLookup(negControlCol + 2) num2str(i + 2)]};%Record location of "Inconclusive" cells to be for yellow highlighting later
            inconclusiveCounter = inconclusiveCounter + 1;
            
        elseif negControlDecision(i,1) == 2 || negControlDecision(i,2) == 2
            
            decisionMat(i,negControlCol) = {''};
            
        else
            
            decisionMat(i,negControlCol) = {'Pass'};
            
            goodCells(goodCounter) = {[columnLookup(negControlCol + 2) num2str(i + 2)]};%Record location of "Pass" cells to be for green highlighting later
            goodCounter = goodCounter + 1;
            
        end
        
        %Positive control - Result is "Pass" if FAM channel amplification
        %(viral RNA) is found, and HEX channel amplification(human RNA) is
        %not
        if posControlDecision(i,1) == 0 || posControlDecision(i,2) == 1
            
            decisionMat(i,posControlCol) = {'Inconclusive'};
            
            inconclusiveCells(inconclusiveCounter) = {[columnLookup(posControlCol + 2) num2str(i + 2)]};%Record location of "Inconclusive" cells to be for yellow highlighting later
            inconclusiveCounter = inconclusiveCounter + 1;
            
        elseif posControlDecision(i,1) == 2 || posControlDecision(i,2) == 2
            
            decisionMat(i,posControlCol) = {''};
            
        else
            
            decisionMat(i,posControlCol) = {'Pass'};
            
            goodCells(goodCounter) = {[columnLookup(posControlCol + 2) num2str(i + 2)]};%Record location of "Pass" cells to be for green highlighting later
            goodCounter = goodCounter + 1;
            
        end
    end

    % Make sample decisions
    sampleCols = 1:nColPlate;
    sampleCols([posControlCol negControlCol]) = [];
        
    for i = 1:length(sampleCols)
               
       sampleWells = sampleCols(i):nColPlate:nWell;
       sampleDecision = 2*ones(nRowPlate,2);
                   
       %Make decisions for each sample channel
       for j = 1:length(sampleWells)
           
           if ~isnan(find(rowNdx == sampleWells(j)))
               
               currRowNdxSample = find(rowNdx == sampleWells(j));
               
               %Place the sample label in the output plate
               if j == 1
                   
                   newPlate(1,2+sampleCols(i)) = raw(begin + currRowNdxSample(1),4); 
                   
               end

               if ~isnan(num(currRowNdxSample(1),CtCol)) && num(currRowNdxSample(1),CtCol) < Ct%FAM channel decision

                    sampleDecision(j,1) = 1;
                    
               else
                   
                   sampleDecision(j,1) = 0;

               end

               if ~isnan(num(currRowNdxSample(2),CtCol)) && num(currRowNdxSample(2),CtCol) < Ct%HEX channel decision

                    sampleDecision(j,2) = 1;
                    
               else
                   
                   sampleDecision(j,2) = 0;

               end
               
           end
           
       end
       
       % Create labels based on sample decisions. HEX (Human) amplification is a
       % requirement for a conclusive decision. When HEX amplification is
       % observed, FAM (virus) amplification determines "+" (amplification) or "-" (no amplification) result.
       for k = 1:nRowPlate
           
           if sampleDecision(k,1) == 1 && sampleDecision(k,2) == 1
               
               decisionMat(k,sampleCols(i)) = {'+'};
               
               badCells(badCounter) = {[columnLookup(sampleCols(i) + 2) num2str(k + 2)]};%Record location of "+" cells to be for red highlighting later
               badCounter = badCounter + 1;
               
           elseif sampleDecision(k,1) == 0 && sampleDecision(k,2) == 1
               
               decisionMat(k,sampleCols(i)) = {'-'};
               
               goodCells(goodCounter) = {[columnLookup(sampleCols(i) + 2) num2str(k + 2)]};%Record location of "-" cells to be for green highlighting later
               goodCounter = goodCounter + 1;
               
           elseif sampleDecision(k,1) == 2 || sampleDecision(k,2) == 2
               
               decisionMat(k,sampleCols(i)) = {''};
                       
           else
               
               decisionMat(k,sampleCols(i)) = {'Inconclusive'};
               
               inconclusiveCells(inconclusiveCounter) = {[columnLookup(sampleCols(i) + 2) num2str(k + 2)]};%Record location of "Inconclusive" cells to be for yellow highlighting later
               inconclusiveCounter = inconclusiveCounter + 1;
               
           end
               
       end

    end
    
    %Determine whether the HEX intensity threshold is set properly. This
    %is necessary because FAM bleedthrough is often observed in the HEX
    %channel. Thus, the HEX intensity threshold must be set to properly
    %filter out this bleedthrough. Practically, the HEX intensity threshold
    %should be set above the max intensity value of the HEX signal in the positive control wells, where no human RNA is present. 
    numPosWells = find(posControlWells == rowNdx);
    posControlWellLabels = posControlWells(1:length(numPosWells)/2);
    hexThresh = 0;
    setHexThresh = num(2,16);%pull the set HEX threshold from the data
    
    %Figure out if there is ever a point in the data where the observed HEX
    %value in the positive control wells is greter than the set HEX
    %threshold
    for i = 1:length(numAmp(begin:end,1))
        
        isHex = rawAmp{begin + i,3};
        
        if sum(numAmp(i,1) == posControlWellLabels) && strcmp(isHex(1:3), 'HEX') && numAmp(i,5) > hexThresh
            
            hexThresh = numAmp(i,5);
            
        end
        
    end
    
    %Warn the user if the HEX threshold is improperly set
    if hexThresh > setHexThresh
        
        warning('%s_Hex Threshold Improperly Set',fileName)
        newPlate(1,1) = {'Hex Threshold Improperly Set'};
        
    end
    
    %Intialize output spreadsheet
    [r,c] = size(decisionMat);
    
    newPlate(3:3+r-1,3:3+c-1) = decisionMat;
    xlswrite(sprintf('%s_Processed.xlsx',fileName(1:dotLoc-1)),newPlate)%write sheet with correct labels. Append "_Processed" to file name
    
    %Reopen excel to edit cell colors based on label
    outName = sprintf('%s%s%s_Processed.xlsx',directoryName,'\',fileName(1:dotLoc-1));
    Excel = actxserver('excel.application');
    WB = Excel.Workbooks.Open(outName);% Get Workbook object
    
    %Add green color to the appropriate cells based on description above
    %(lines 24-29)
    if goodCounter > 1
        
        for i = 1:goodCounter - 1 
        
        WB.Worksheets.Item(1).Range(goodCells{i}).Interior.Color = hex2dec('00FF00');
        
        end
        
    end
    
    %Add red color to the appropriate cells based on description above
    %(lines 24-29)
    if badCounter > 1
        
        for i = 1:badCounter - 1 
        
        WB.Worksheets.Item(1).Range(badCells{i}).Interior.Color = hex2dec('0000FF');
        
        end
        
    end
    
    %Add yellow color to the appropriate cells based on description above
    %(lines 24-29)
    if inconclusiveCounter > 1
        
        for i = 1:inconclusiveCounter - 1 
        
        WB.Worksheets.Item(1).Range(inconclusiveCells{i}).Interior.Color = hex2dec('00FFFF');
        
        end
        
    end
    
    %Close Excel
    WB.Save();% Save Workbook
    WB.Close();% Close Workbook
    Excel.Quit();% Quit Excel
    
end