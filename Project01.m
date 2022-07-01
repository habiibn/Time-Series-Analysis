clc;
clear;

%% Set up the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 9);

% Specify sheet and range
opts.Sheet = "Sheet1";
opts.DataRange = "A2:I262";

% Specify column names and types
opts.VariableNames = ["Periods", "AverageOfO3Mean", "AverageOfCOMean", "AverageOfSO2Mean", "AverageOfNO2Mean", "AverageOfO3AQI", "AverageOfCOAQI", "AverageOfSO2AQI", "AverageOfNO2AQI"];
opts.VariableTypes = ["string", "double", "double", "double", "double", "double", "double", "double", "double"];

% Specify variable properties
opts = setvaropts(opts, "Periods", "WhitespaceRule", "preserve");
opts = setvaropts(opts, "Periods", "EmptyFieldRule", "auto");

% Import the data
Airpollution = readtable("E:\Semester 6\Pengolahan Sinyal Pengaturan\Tugas\Proyek 1\Airpollution.xlsx", opts, "UseExcel", false);


%% Clear temporary variables
clear opts

%% Plot
x = table2array(Airpollution(:,1));
y = table2array(Airpollution(:,9));
X = datetime(x, 'InputFormat', 'MMM-yyyy', 'Format', 'MMM yyy');
% plot(y)

%% Uji stasioner terhadap varians
[~, lambda0] = boxcox(y);
[y_boxcox, ~] = boxcox(y);

[~, lambda1] = boxcox(y_boxcox);
[y_boxcox1, ~] = boxcox(y_boxcox);
%% Uji stasioner terhadap mean
y_mean = movmean(y,12);
%% Differencing
y_diff = diff(y);
y_diff2 = diff(y_diff);
y_diff3 = diff(y_diff2);

%% Model prediksi
Mdl = arima('ARLags',1,'Constant',0,'D',1,'Seasonality',12,'MALags',1,'SMALags',12);
EstMdl = estimate(Mdl,y);
residuals = infer(EstMdl,y);
prediction = y+residuals;
%% Plotting Model
figure
h1 = plot(X,y,'Color',[0,0,0]);
hold on
h2 = plot(X,y_mean,'b--');
h3 = plot(X,prediction,'r:');
legend([h1 h2 h3],'Data Asli','Mean Data', 'Arima Model');
hold off

%% Plotting 100 Data 
% figure
% h1 = plot(X(1:100),y(1:100),'b-');
% hold on
% h2 = plot(X(1:100),y_mean(1:100),'Color',[0.7,0.7,0.7]);
% h3 = plot(X(1:100),prediction(1:100),'m-');
% legend([h1 h2 h3],'Data Asli','Mean Data', 'sArima Model');
% title('sArima(2,1,1) untuk SMA(12)');
% hold off

%% MSE, AIC, BIC
y_mse = mean((y - prediction).^2);

[EstMdl0,EstParamCov0,logL0,info0]=estimate(arima('Constant',0,'D',1,'Seasonality',12,'MALags',1,'SMALags',12),y_diff);
[EstMdl1,EstParamCov1,logL1,info1]=estimate(arima('ARLags',1,'Constant',0,'D',1,'Seasonality',12,'MALags',1,'SMALags',12),y_diff);
[EstMdl2,EstParamCov2,logL2,info2]=estimate(arima('ARLags',2,'Constant',0,'D',1,'Seasonality',12,'MALags',1,'SMALags',12),y_diff);

[aic,bic] = aicbic([logL0,logL1,logL2],[2,3,4]);
%% Identifikasi Model
% Mdl = arima('AR',{0.5,-0.3},'MA',0.5,'SMA',0.5,'SMALags',12,'Constant',0,'Variance',4);
% rng(200);
% Y = simulate(Mdl,261);

% MdlTemplate = arima('AR',[0.5 -0.3],'MA',0.8);
% EstMdl = estimate(MdlTemplate,y(1:length(y)-30));
% % [E,V] = infer(MdlTemplate,y);
% 
% [YF,YMSE] = forecast(EstMdl,30,'Y0',y(1:length(y)-30));
% YS = simulate(MdlTemplate);
%%
%h2 = plot(y_mean,'Color','g');
% h3 = plot(261:290,YF,'b','LineWidth',2);
% h4 = plot(261:290,YF + 1.96*sqrt(YMSE),'r:','LineWidth',2);
%      plot(261:290,YF - 1.96*sqrt(YMSE),'r:','LineWidth',2);
% legend([h1],'Observed','Forecast','95% Confidence Interval','Location','NorthWest');
% title('30-Period Forecasts and Approximate 95% Confidence Intervals');

%%
% Mdl = arima('D',1,'Seasonality',12,'MALags',1,'SMALags',12);
% EstMdl = estimate(Mdl,y(1:length(y)-12));
% [Y,YMSE,V] = forecast(EstMdl,12,'Y0',y(1:length(y)-12));
% ForecastInt = [Y,Y] + 1.96*[-sqrt(YMSE), sqrt(YMSE)];
% 
% %% Forecasting
% figure
% h1 = plot(y);
% title('{\bf Forecasted Monthly NO2 Rate Average}')
% hold on
% h2 = plot(Y,'Color','r','LineWidth',2);
% h3 = plot(ForecastInt,'k--','LineWidth',2);
% datetick
% legend([h1,h2,h3(1)],'Observations','MMSE Forecasts',...
%     '95% MMSE Forecast Intervals','Location','NorthWest')
% axis tight