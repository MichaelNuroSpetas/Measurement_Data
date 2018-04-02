% Script to import meassurment data
%% Import sheet 1
[~, ~, sheet1] = xlsread('New Microsoft Excel Worksheet.xlsx','Sheet1','A1:Z18726');
sheet1(cellfun(@(x) ~isempty(x) && isnumeric(x) && isnan(x),sheet1)) = {''};

idx = cellfun(@ischar, sheet1);
sheet1(idx) = cellfun(@(x) string(x), sheet1(idx), 'UniformOutput', false);

%% Clear temporary variables
clearvars idx;

%% Import sheet 2
[~, ~, sheet2] = xlsread('New Microsoft Excel Worksheet.xlsx','Sheet2','A1:D18726');
sheet2(cellfun(@(x) ~isempty(x) && isnumeric(x) && isnan(x),sheet2)) = {''};

idx = cellfun(@ischar, sheet2);
sheet2(idx) = cellfun(@(x) string(x), sheet2(idx), 'UniformOutput', false);

%% Clear temporary variables
clearvars idx;

%% Import sheet 3
[~, ~, sheet3] = xlsread('New Microsoft Excel Worksheet.xlsx','Sheet3','A1:AA1878');

sheet3(cellfun(@(x) ~isempty(x) && isnumeric(x) && isnan(x),sheet3)) = {''};

idx = cellfun(@ischar, sheet3);
sheet3(idx) = cellfun(@(x) string(x), sheet3(idx), 'UniformOutput', false);

%% Clear temporary variables
clearvars idx;

%% Import updated Vækerø(P42_FT01) and IPU_FT03 data
[~, ~, sheet4] = xlsread('veas_data_010318.xlsx','veas_data_010318','A7:K18732');
sheet4(cellfun(@(x) ~isempty(x) && isnumeric(x) && isnan(x),sheet4)) = {''};

idx = cellfun(@ischar, sheet4);
sheet4(idx) = cellfun(@(x) string(x), sheet4(idx), 'UniformOutput', false);


%% Merge sheet 1 (trimmed) sheet 2 horizontally together
sheet2t = sheet2(:, 2:end);
sheet12 = horzcat(sheet1, sheet2t);

%% Trimming and resampling sheet 3
% Trimming text on top to be able to make into a double matrix
sheet3t = sheet3(7:end, 2:end);

msheet3td = cell2mat(sheet3t);       %Convert cell to double matrix
msheet3r = resample(msheet3td,10,1);  %Resampled sheet with filter set to #, default filter coefficient is 10
%msheet3i = interp1((1:1872),msheet3td',(1:18720));  %Resampled sheet with filter set to 0, default filter coefficient is 10
sheet3rc = num2cell(msheet3r);       %Resampled double sheet to cell
sheet3toptext = sheet3(1:6,2:end);  %Top text cell made from sheet 3
sheet3rnt = vertcat(sheet3toptext, sheet3rc); %resampled sheet 3 merged with top text

%% Merge sheet 12 and 3 horizontally together
sheet123 = horzcat(sheet12, sheet3rnt);

% Merge sheet123 and a (trimmed) sheet4 horizontally together
sheet1234 = horzcat(sheet123, sheet4(:,2),sheet4(:,10));

%% Trim top text, time and INVALID data indicated by # down to columns 580-585 in the start, then convert to 1 large double matrix for data analysis, modification & plotting
%msheet1234t = cell2mat(sheet1234(7:end,2:end));
msheet1234t = cell2mat(sheet1234(587:end,2:end));

%% Seperate matrix into induvidually named arrays 
% PASL are inflow measurements
PASL1_SJ = msheet1234t(:,29);
PASL2_ELJ = msheet1234t(:,30);
PASL3_NI = msheet1234t(:,31);
PASL4_STV = msheet1234t(:,32);
PASL5_OS = msheet1234t(:,33);
PASL6_BL = msheet1234t(:,34);
PASL7_LE = msheet1234t(:,35);
PASL8_HAG = msheet1234t(:,36);
PASL9_SKU = msheet1234t(:,37); % This sensor is not visible on tegning nr. 1235_001
PASL10_HO = msheet1234t(:,38);
PASL11_BI = msheet1234t(:,39);
PASL12_SAV1 = msheet1234t(:,40); % There two measurements for Sandvika Vest 
PASL13_SAV2 = msheet1234t(:,41); % Both appear to have functioning data
% PASL135_SAV3 = msheet1234t(:,42); % Useless null data
PASL14_BJ = msheet1234t(:,43);
PASL15_HAM = msheet1234t(:,44);
PASL16_SAO = msheet1234t(:,45);
PASL17_SKY = msheet1234t(:,46); %% There is a SKY1 at 47, but this appears to be invalid data
%PASL18_EV = msheet1234t(:,?); % This one is missing
PASL19_FO = msheet1234t(:,48);
PASL20_SO = msheet1234t(:,49);
PASL21_SKA = msheet1234t(:,50); % The order of Stabekk and Skallum appear to be switched in one of the documents
PASL22_STB = msheet1234t(:,51);
PASL23_JA = msheet1234t(:,52);
PASL24_LI = msheet1234t(:,53);
PASL25_VA = msheet1234t(:,54); %Something wrong with the PASL_VA
FT_VA = msheet1234t(:,55); % The new data set provided when the VA flow measurement was forgotten
%PA1623_FT16 = msheet1234t(:,17); % Frognerparken
% Missing data from Lilleaker and up will not be supplied

%% Level measurements
IPU_LT = msheet1234t(:,3); %Level of IPU basin in meters
EV_LT = msheet1234t(:,4); %Level inside the Engervann tunnel
PA1623_LT07 = msheet1234t(:,6); %Level inside the Torshov tunnel (Majorstua - Fagerlia)
PA1623_LT01 = msheet1234t(:,7); % Level inside "Sentrumtunnelen/Festingstunnelen"
PA1623_LT02 = msheet1234t(:,8); % Level inside "Sentrumtunnelen/Festingstunnelen"
PA1623_LT03 = msheet1234t(:,9); % Level inside "Sentrumtunnelen/Festingstunnelen"
PA1623_LT04 = msheet1234t(:,10); % Level inside "Sentrumtunnelen/Festingstunnelen" (Frognerparken)

%% Pump flow
% Flow from pumps to Frognerparken
PA1623_FT = sum(msheet1234t(:,11:16),2);
% OR use the already summated column, however there are differences.
PA1623_FT16 = msheet1234t(:,17);

%% Grouped inflows
% Inflow to the VEAS-Engervann basin
PASL1_16_VEAS_EV = msheet1234t(:,29:46);
SUM_PASL_VEAS_EV = sum(PASL1_16_VEAS_EV,2); % Sum of all inflow values (rows) to the VEAS_EV basin
% Inflow to the Engervann-Vækerø basin
PASL17_24_EV_VA = msheet1234t(:,48:53); %Note EV is missing
SUM_PASL_EV_VA = sum(PASL17_24_EV_VA,2); % Sum of all inflow values (rows) to the EV_VA basin

% Inflow to IPU basin
IPU_FT_03missing = sum(msheet1234t(:,20:26),2); %IPU_FB03 flowdata was later supplied for a complete sum
IPU_FT = sum(horzcat(IPU_FT_03missing,msheet1234t(:,56)),2); % Total IPU_FT
% Wash water flow to IPU
TSP_FB07 = sum(msheet1234t(:,27:28),2);
% Direct flow to IPU ("no delay")
AL_FT10 = msheet1234t(:,18);
AAR_FT01 = msheet1234t(:,19);
% All direct inflow to IPU with no delay
ALL_DIRECT_IPU = horzcat(TSP_FB07,AAR_FT01,AL_FT10);
SUM_DIRECT_IPU = sum(ALL_DIRECT_IPU,2);
% Export all matrices containing the word SUM for the grouped flows

%% Write excel file containing all nessecary flows

%xlswrite('SUM_DATASETS2', [IPU_FT, SUM_DIRECT_IPU, SUM_PASL_VEAS_EV, SUM_PASL_EV_VA, FT_VA]);

%% Test plotting
td = 1:18720;
%plot(td, (msheet123t(:,21:26)))
% Compare original data with resampled data
%interp1_8 = (interp1((1:1872),msheet3td(:,8),td))'; %  NaN data error
figure(1)
plot(((1:1872)*10),msheet3td(:,8), '-r',td,msheet3r(:,8),'-b'), grid % Comparing resampled data to non-resampled data
%plot(td, msheet3r(:,8),td, msheet3i(:,8))
grid minor
legend('Orginal data sampled at 10min intervals', 'Data resampled at 1min intervals');
%% Plotting Vækerø vs ipu flow sum


tdsp = 180; %timedisplacement, transportating time
tvec = 1:(length(FT_VA)-tdsp);
scale = (ones(1,length(FT_VA))*1200)';
VA_scaled = sum(horzcat(scale,FT_VA),2);
plot(tvec, VA_scaled(1:(end-tdsp),:), tvec, IPU_FT((tdsp+1):end,:), tvec, PA1623_FT(1:(end-tdsp),:))
legend('Vækerø flow measurement', 'IPU flow', 'Pump flow to Frognerparken')