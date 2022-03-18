%Final Project

data = xlsread('Final Project v3.xlsx','Returns','B2:U13')
%DATA = xlsread('file_name','sheet_name','range')

[~,names] = xlsread('Final Project v3.xlsx','Returns','B1:U1')

%Part A

C = cov(data) %covariance matrix
R = mean(data) %expected returns

%Part B

p = Portfolio;
p = setAssetMoments(p, R, C);
p = setDefaultConstraints(p);


p = setAssetList(p,names);

plotFrontier(p);
xlabel('Risk (Std Dev of Return)');
ylabel('Expected Returns');


%Part C

avg_return=mean(R)


opt_portfolio = estimateFrontierByReturn(p, avg_return)
%first column of opt_portfolio matrix gives you the solution for the
%problem of target returns

opt_portfolio
opt_portfolio(:,1)

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Part D

%Portfolio Optimization Against a Benchmark
benchPrice = xlsread('Final Project v3.xlsx','Price','B2:b14');
assetPrice = xlsread('Final Project v3.xlsx','Price','c2:w14');
[~,assetNames] = xlsread('Final Project v3.xlsx','Price','c1:w1');
[~,benchName] = xlsread('Final Project v3.xlsx','Price','b1');
Dates=xlsread('Final Project v3.xlsx','Price','a2:a14');

%%convert the dates, they look wierd in matlab

d = datetime(Dates, 'ConvertFrom', 'excel')

assetP = assetPrice./assetPrice(1, :);  
benchmarkP = benchPrice / benchPrice(1);


figure;
plot(d,assetP);
hold on;
plot(d,benchmarkP,'LineWidth',3,'Color','k');
hold off;
xlabel('Months');
ylabel('Prices');
title('Asset Prices and Benchmark');
grid on;


%Compute Returns and Risk-Adjusted Returns
benchReturn = tick2ret(benchPrice);
assetReturn = tick2ret(assetPrice);

benchRetn = mean(benchReturn);
benchRisk =  std(benchReturn);
assetRetn = mean(assetReturn);
assetRisk =  std(assetReturn);

scale = 21;

assetRiskR = sqrt(scale) * assetRisk;
benchRiskR = sqrt(scale) * benchRisk;
assetReturnR = scale * assetRetn;
benchReturnR = scale * benchRetn;

figure;
scatter(assetRiskR, assetReturnR, 6, 'm', 'Filled');
hold on
scatter(benchRiskR, benchReturnR, 6, 'g', 'Filled');
for k = 1:length(assetNames)
    text(assetRiskR(k) + 0.005, assetReturnR(k), assetNames{k}, 'FontSize', 8);
end
text(benchRiskR + 0.005, benchReturnR, 'Benchmark', 'Fontsize', 8);
hold off;

xlabel('Risk (Std Dev of Return)');
ylabel('Expected Annual Return');
grid on;

%Set Up a Portfolio Optimization
p = Portfolio('AssetList',assetNames);
p = setDefaultConstraints(p);
activReturn = assetReturn - benchReturn;
pAct = estimateAssetMoments(p,activReturn,'missingdata',false);

%Compute the Efficient Frontier Using the Portfolio Object
pwgtAct = estimateFrontier(pAct, 20); % Estimate the weights.
[portRiskAct, portRetnAct] = estimatePortMoments(pAct, pwgtAct); % Get the risk and return.

% Extract the asset moments and names.
[assetActRetnMonthly, assetActCovarMonthly] = getAssetMoments(pAct);
assetActRiskMonthly = sqrt(diag(assetActCovarMonthly));
assetNames = pAct.AssetList;


% Rescale.
assetActRiskAnnual = sqrt(scale) * assetActRiskMonthly;
portRiskAnnual  = sqrt(scale) *  portRiskAct;
assetActRetnAnnual = scale * assetActRetnMonthly;
portRetnAnnual = scale *  portRetnAct;

figure;
subplot(2,1,1);
plot(portRiskAnnual, portRetnAnnual, 'bo-', 'MarkerFaceColor', 'b');
hold on;

scatter(assetActRiskAnnual, assetActRetnAnnual, 12, 'm', 'Filled');
hold on;
for k = 1:length(assetNames)
    text(assetActRiskAnnual(k) + 0.005, assetActRetnAnnual(k), assetNames{k}, 'FontSize', 8);
end

hold off;

xlabel('Risk (Std Dev of Active Return)');
ylabel('Expected Active Return');
grid on;

subplot(2,1,2);
plot(portRiskAnnual, portRetnAnnual./portRiskAnnual, 'bo-', 'MarkerFaceColor', 'b');
xlabel('Risk (Std Dev of Active Return)');
ylabel('Information Ratio');
grid on;

%Perform Information Ratio Maximization Using Optimization Toolbox™
%The infoRatioTargetReturn local function is called as an objective function in an optimization routine (fminbnd) 
%that seeks to find the target return that maximizes the information ratio and minimizes a negative information ratio.

objFun = @(targetReturn) -infoRatioTargetReturn(targetReturn,pAct);
options = optimset('TolX',1.0e-8);
[optPortRetn, ~, exitflag] = fminbnd(objFun,min(portRetnAct),max(portRetnAct),options);

%Get weights, information ratio, and risk return for the optimal portfolio.
[optInfoRatio,optWts] = infoRatioTargetReturn(optPortRetn,pAct);
optPortRisk = estimatePortRisk(pAct,optWts) 

%Plot the Optimal Portfolio
%Verify that the portfolio found is indeed the maximum information-ratio portfolio.
% Rescale.
optPortRiskAnnual = sqrt(scale) * optPortRisk;
optPortReturnAnnual = scale * optPortRetn;

figure;
subplot(2,1,1);

scatter(assetActRiskAnnual, assetActRetnAnnual, 6, 'm', 'Filled');
hold on
for k = 1:length(assetNames)
    text(assetActRiskAnnual(k) + 0.005,assetActRetnAnnual(k),assetNames{k},'FontSize',8);
end
plot(portRiskAnnual,portRetnAnnual,'bo-','MarkerSize',4,'MarkerFaceColor','b');
plot(optPortRiskAnnual,optPortReturnAnnual,'ro-','MarkerFaceColor','r');
hold off;

xlabel('Risk (Std Dev of Active Return)');
ylabel('Expected Active Return');
grid on;

subplot(2,1,2);
plot(portRiskAnnual,portRetnAnnual./portRiskAnnual,'bo-','MarkerSize',4,'MarkerFaceColor','b');
hold on
plot(optPortRiskAnnual,optPortReturnAnnual./optPortRiskAnnual,'ro-','MarkerFaceColor','r');
hold off;

xlabel('Risk (Std Dev of Active Return)');
ylabel('Information Ratio');
title('Information Ratio with Optimal Portfolio');
grid on;

%Display the Portfolio Optimization Solution
assetIndx = optWts > .001;
results = table(assetNames(assetIndx)', optWts(assetIndx)*100, 'VariableNames',{'Asset', 'Weight'});
disp('Maximum Information Ratio Portfolio:')
%Maximum Information Ratio Portfolio:
disp(results)



