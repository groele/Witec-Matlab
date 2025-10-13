% -------------------------------------------------------------------------
% Spectral Data Analysis Script
% 功能：
%   本脚本用于对光谱数据进行批量分析和处理，支持 Excel、CSV、TXT 等格式。
%   主要功能包括：
%     1. 选择单个数据文件并读取数据；
%     2. 按指定起始行和终止行截取数据；
%     3. 去除基线值并对每列数据进行平滑处理（loess, lowess, movmean, sgolay）；
%     4. 绘制原始数据与平滑数据的叠加图，并标注峰值及 FWHM；
%     5. 计算面积（积分）及峰值特征，并生成汇总表格；
%     6. 绘制特征随功率或其他变量变化的趋势图；
%     7. 将原始数据、平滑数据、分析结果和参数日志保存到指定结果文件夹。
%
% 使用说明：
%   1. 修改脚本中 Start_Row、End_Row、baseline_value、smooth_type 等参数以适应数据；
%   2. 运行脚本，选择数据文件；
%   3. 程序将自动生成以“日期_文件名”为名的结果文件夹，并保存结果及图像；
%   4. 脚本中辅助函数已实现平滑、峰值标注、FWHM计算、积分等功能。
%
% 支持文件格式：
%   - Excel: .xls, .xlsx
%   - 文本/CSV: .txt, .csv
%
% 注意事项：
%   - 数据文件第一行应为列标题；
%   - 数据文件第一列为 X 轴（能量、波数等），后续列为 Y 值；
%   - 平滑参数需根据实际数据调整，以保证平滑效果与特征保真度。
%
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------

clc; clear; close all;
tic  % 启动计时器

%% -----------------------------
% 用户参数设置
%% -----------------------------
startRow       = 600;      % 数据起始行
endRow         = 1110;     % 数据终止行
baselineValue  = 486;      % 基线值

% baselineValue  = 0;      % 基线值
smoothType     = 'loess';  % 平滑方法：'loess','lowess','movmean','sgolay'
smoothParam    = 0.05;     % loess 的跨度
savePlot       = true;     % 是否保存图像

%% -----------------------------
% 文件选择
%% -----------------------------
[filename, pathname] = uigetfile({'*.xls;*.xlsx;*.txt;*.csv','文本文件 (*.xls;*.xlsx;*.txt;*.csv)'; ...
                                  '*.*','所有文件 (*.*)'}, ...
                                  '请选择一个数据文件');
if isequal(filename,0); return; end
Filepath = fullfile(pathname,filename);

%% -----------------------------
% 创建结果目录
%% -----------------------------
[~, basename, ~] = fileparts(filename);
dateStr  = char(datetime('now','Format','yyyyMMdd'));
folderTag = [dateStr '_' basename];
targetDir = fullfile(pathname, folderTag);
if exist(targetDir,'dir'); rmdir(targetDir,'s'); end
mkdir(targetDir);

%% -----------------------------
% 数据读取与预处理
%% -----------------------------
Data = readcell(Filepath);  % 支持混合数据
rawHead = Data(1,2:end);    % 去掉第一列
% 转为数值，若无法转为数值保持原字符串
rowTit = cellfun(@(x) tryStr2Num(x), rawHead,'Uni',false);  

% 提取 X + 每列 Y
XY = cell2mat([Data(2:end,1), Data(2:end,2:end)]);  % 所有列都保留
XY = XY(startRow:endRow,:);                        % 截取行
energyX = XY(:,1);
yRaw    = XY(:,2:end) - baselineValue;

% 检查列数一致性
if length(rowTit) ~= size(yRaw,2)
    error('列标题数 (%d) 与数据列数 (%d) 不匹配，请检查输入文件！', length(rowTit), size(yRaw,2));
end

%% -----------------------------
% 平滑处理
%% -----------------------------
yFilt = smoothMatrix(yRaw, smoothType, smoothParam);

%% -----------------------------
% 绘图与指标提取
%% -----------------------------
nCol   = size(yRaw,2);
summary = cell(nCol+1,9);
summary(1,:) = {'Power/Temper','X value','Y value','FWHM', ...
                'X value_filtered','Y value_filtered','FWHM_filtered', ...
                'Area','Area_filtered'};

[rPlot,cPlot] = subPlotLayout(nCol);
hFig = figure('Name','Per‑Column Analysis','WindowState','max'); clf

for i = 1:nCol
    subplot(rPlot,cPlot,i); hold on
    
    % 面积图
    area(energyX, yFilt(:,i), 'FaceAlpha',0.18, ...
         'FaceColor',[1 0.4 0.4], 'EdgeColor','none');
    
    plot(energyX, yRaw(:,i),  'b','LineWidth',1.1);
    plot(energyX, yFilt(:,i), 'r','LineWidth',1.1);
    title(rowTit{i});
    
    % 峰值与 FWHM
    [xR,yR] = annotateMaxValue(yRaw,  energyX,i,'b*',[-0.2,0.1]);
    fwhmR  = annotateFWHM(yRaw(:,i),  energyX);
    [xF,yF] = annotateMaxValue(yFilt, energyX,i,'r*',[-0.2,0.1]);
    fwhmF  = annotateFWHM(yFilt(:,i), energyX);
    
    % 面积
    aR = safeTrapz(energyX, yRaw(:,i));
    aF = safeTrapz(energyX, yFilt(:,i));
    
    summary(i+1,:) = {rowTit{i}, xR,yR,fwhmR, xF,yF,fwhmF, aR, aF};
    hold off
end

sgtitle('Power Dependence','FontSize',14);
if savePlot
    saveas(hFig, fullfile(targetDir,['1、Plots_' basename '.png']));
end

%% -----------------------------
% 汇总特征图
%% -----------------------------
powerVals = cell2mat(summary(2:end,1));
energyVal = cell2mat(summary(2:end,5));
peakVal   = cell2mat(summary(2:end,6));
areaVal   = cell2mat(summary(2:end,9));

hFig2 = figure('Name','Summary Plots','WindowState','max');
subplot(1,3,1); plot(powerVals, energyVal,'-o'); xlabel('Power'); ylabel('Energy (eV)');
subplot(1,3,2); plot(powerVals, peakVal,  '-s'); xlabel('Power'); ylabel('Peak Intensity');
subplot(1,3,3); plot(powerVals, areaVal,  '-^'); xlabel('Power'); ylabel('Area (filtered)');
if savePlot
    saveas(hFig2, fullfile(targetDir,['2、Summary_' basename '.png']));
end

%% -----------------------------
% 数据保存
%% -----------------------------
Data = cellfun(@(x) convertMissing(x), Data,'Uni',false);
writecell(Data, fullfile(targetDir,['3、Raw_'      basename '.xlsx']));
writecell([{'Energy_X'}, rowTit; num2cell([energyX, yFilt])], ...
          fullfile(targetDir,['4、Filtered_' basename '.xlsx']));
writecell(summary, fullfile(targetDir,['5、Result_' basename '.xlsx']));

%% -----------------------------
% 参数日志
%% -----------------------------
fid = fopen(fullfile(targetDir,['0、Parameters_' basename '.txt']),'w');
dt = datetime('now', 'Format', 'yyyy-MM-dd HH:mm:ss');
fprintf(fid, 'Processing date            : %s\n',  char(dt));
fprintf(fid,'Data range             : %d → %d\n', startRow, endRow);
fprintf(fid,'X axis range           : %.4f → %.4f\n', energyX(1), energyX(end));
fprintf(fid,'Baseline value         : %.4f\n', baselineValue);
fprintf(fid,'Smoothing              : %s (%.3f)\n', smoothType, smoothParam);
fclose(fid);

toc

%% ------------------------------------------------------------------------
%% 辅助函数
%% ------------------------------------------------------------------------

function y = convertMissing(x)
    if ismissing(x); y=''; else; y=x; end
end

function r = ifelse(c,a,b)
    if c; r=a; else; r=b; end
end

function n = tryStr2Num(x)
    n = str2double(x);
    if isnan(n); n=x; end
end

function S = smoothMatrix(M, type, varargin)
    [m,n] = size(M); S = zeros(m,n);
    for k = 1:n
        v = M(:,k);
        switch lower(type)
            case {'lowess','loess'}
                span = 0.2; if ~isempty(varargin); span=varargin{1}; end
                S(:,k)=smooth(v,span,type);
            case 'movmean'
                w = 5; if ~isempty(varargin); w=varargin{1}; end
                S(:,k)=smoothdata(v,'movmean',w,'omitnan');
            case 'sgolay'
                o=2; f=5;
                if length(varargin)>=2; o=varargin{1}; f=varargin{2}; end
                S(:,k)=smoothdata(v,'sgolay',f,o,'omitnan');
            otherwise
                error('Unknown smooth type');
        end
    end
end

function [mxX,mxY]=annotateMaxValue(M,x,i,marker,off)
    if nargin<5; off=[0,0]; end
    if nargin<4; marker='ro'; end
    [mxY,id] = max(M(:,i)); mxX=x(id);
    plot(mxX,mxY,marker,'MarkerSize',7,'LineWidth',1.2);
    text(mxX+off(1), mxY+off(2), sprintf('(%.2f, %.2f)',mxX,mxY), ...
         'Vert','bottom','Horiz','right');
end

function F=annotateFWHM(y,x,style)
    if nargin<3; style='g--'; end
    h = y >= max(y)/2;
    if ~any(h); F=NaN; return; end
    idx = h;
    F = abs(x(find(idx,1,'last')) - x(find(idx,1,'first')));
    xx = x(idx);
    plot(xx, repmat(max(y)/2,1,numel(xx)), style, 'LineWidth',1);
    text(xx(1), max(y)/2, sprintf('FWHM=%.2f', F), ...
         'Vert','bottom','Horiz','left','Color','green');
end

function [r,c]=subPlotLayout(n)
    r = ceil(sqrt(n));
    c = ceil(n/r);
end

function A=safeTrapz(x,y)
    if x(1)>x(end)
        A = trapz(flip(x), flip(y));
    else
        A = trapz(x,y);
    end
    A = abs(A);
end
