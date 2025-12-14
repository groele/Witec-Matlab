% =========================================================================
% Spectral Data Batch Analysis (Unified Framework)
% 功能：
%   - 从单个数据文件（Excel/CSV/TXT）读取光谱数据
%   - 按 Excel 行号截取指定行范围
%   - 对每列光谱进行基线校正与平滑
%   - 计算峰值、FWHM、积分面积（Raw & Smooth）
%   - 绘制单列光谱叠加图与趋势图（Peak Energy / Intensity / Area）
%   - 导出原始/滤波数据、结果 summary，以及 Publication-Ready CSV
%
% 作者：Shikun Hou
% 版本：Unified Framework v1.0
% 更新时间：2025-11-26
% =========================================================================

clc; clear; close all;
tic;

%% ========================================================================
% 1. 用户参数设置（Parameters）
%% ========================================================================
params = struct();

% 注意：row_start / row_end 为 Excel 中的"绝对行号"，第 1 行是标题行
% 数据区从第 2 行开始，所以这里填写的是 Excel 原始行号
params.row_start    = 400;       % 数据起始行（Excel 行号）
params.row_end      = 1300;      % 数据终止行（Excel 行号）

params.baseline     = 480;       % 基线值
params.smooth_type  = 'loess';   % 'loess','lowess','movmean','sgolay'
params.smooth_param = 0.00002;   % loess/lowess: 相对跨度; movmean/sgolay: 点数或比例

params.save_plot    = true;      % 是否保存图像
params.show_figure  = true;      % 是否在屏幕显示图像

%% ========================================================================
% 2. 文件选择与输出目录设置（IO）
%% ========================================================================
[file_path, target_dir, base_name] = setupIO_Batch();

%% ========================================================================
% 3. 数据读取与预处理（Import & Preprocess）
%% ========================================================================
[data_x, data_y_raw, headers] = importData_Batch(file_path, params);

%% ========================================================================
% 4. 数据平滑（Smoothing）
%% ========================================================================
data_y_filt = smoothMatrixGeneric(data_y_raw, params.smooth_type, params.smooth_param);

%% ========================================================================
% 5. 单列分析与光谱绘制（Peak, FWHM, Area）
%% ========================================================================
summary = analyzeSpectra_Batch( ...
    data_x, data_y_raw, data_y_filt, headers, target_dir, base_name, params);

%% ========================================================================
% 6. 趋势图绘制（Summary Plots）
%% ========================================================================
plotSummary_Batch(summary, target_dir, base_name, params);

%% ========================================================================
% 7. 论文绘图友好数据导出（Publication-Ready Data）
%% ========================================================================
exportPublicationReady_Batch(summary, target_dir);

%% ========================================================================
% 8. 保存原始/滤波数据、summary 与参数日志（Save All Output）
%% ========================================================================
saveAllOutput_Batch(summary, data_x, data_y_filt, headers, target_dir, base_name, params);

fprintf('Spectral Batch 分析完成，总耗时：%.2f 秒\n', toc);

%% ========================================================================
% 9. 本脚本使用的本地函数
%% ========================================================================

function [file_path, target_dir, base_name] = setupIO_Batch()
% -------------------------------------------------------------------------
% setupIO_Batch
% 功能：选择数据文件并创建输出目录
% -------------------------------------------------------------------------
[filename, pathname] = uigetfile( ...
    {'*.xls;*.xlsx;*.txt;*.csv','文本/表格文件 (*.xls;*.xlsx;*.txt;*.csv)'; ...
     '*.*','所有文件 (*.*)'}, ...
     '请选择光谱数据文件');

if isequal(filename,0)
    error('未选择文件，脚本终止。');
end

file_path = fullfile(pathname, filename);
[~, base_name, ~] = fileparts(filename);

date_str   = char(datetime('now','Format','yyyyMMdd'));
folder_tag = [date_str '_' base_name];
target_dir = fullfile(pathname, folder_tag);

if exist(target_dir,'dir')
    rmdir(target_dir,'s');
end
mkdir(target_dir);
end

function [data_x, data_y_raw, headers] = importData_Batch(file_path, params)
% -------------------------------------------------------------------------
% importData_Batch
% 功能：
%   - 读取光谱文件
%   - 第 1 行为标题行：第 2 列开始为功率/温度/条件标签
%   - 第 2 行起为数据区：第一列为 X（如 Energy），后续列为光谱
%   - 使用 Excel 绝对行号指定截取范围
% 输出：
%   - data_x      : X 轴列向量
%   - data_y_raw  : 原始 Y 矩阵（已减基线）
%   - headers     : 每列 Y 对应的标题（功率/条件）
% -------------------------------------------------------------------------
Data = readcell(file_path);
n_rows = size(Data,1);

if params.row_start < 2 || params.row_end > n_rows || params.row_start >= params.row_end
    error('行号设置不合法：row_start=%d, row_end=%d, 总行数=%d', ...
        params.row_start, params.row_end, n_rows);
end

raw_head = Data(1, 2:end);                        % 第 1 行的列标题，跳过第一列
headers  = cellfun(@(x) tryStr2Num(x), raw_head, 'UniformOutput', false);

% 数据区为 Data(2:end,:)，对应 Excel 的第 2 行到第 n 行
XY_all = cell2mat([Data(2:end,1), Data(2:end,2:end)]);

% 将 Excel 行号转换为数据区索引：行2→索引1
idx_start = params.row_start - 1;
idx_end   = params.row_end   - 1;

n_data_rows = size(XY_all,1);
if idx_start < 1 || idx_end > n_data_rows || idx_start >= idx_end
    error('转换后的数据索引不合法：idx_start=%d, idx_end=%d, 数据区行数=%d', ...
        idx_start, idx_end, n_data_rows);
end

XY = XY_all(idx_start:idx_end, :);

data_x      = XY(:,1);
data_y_raw  = XY(:,2:end) - params.baseline;

if numel(headers) ~= size(data_y_raw,2)
    error('列标题数 (%d) 与数据列数 (%d) 不匹配！', ...
        numel(headers), size(data_y_raw,2));
end

data_x = data_x(:);
end

function y = tryStr2Num(x)
% -------------------------------------------------------------------------
% tryStr2Num
% 功能：尽量将字符串转换为数值；否则保持原类型
% -------------------------------------------------------------------------
if isnumeric(x)
    y = x;
    return;
end
v = str2double(x);
if isnan(v)
    y = x;
else
    y = v;
end
end

function S = smoothMatrixGeneric(M, method, param)
% -------------------------------------------------------------------------
% smoothMatrixGeneric
% 功能：对矩阵 M 的每一列按指定方法进行平滑
%   method:
%     - 'loess','lowess','rloess','rlowess' : param 为相对跨度（0~1）
%     - 'movmean'                           : param 为窗口点数或比例
%     - 'sgolay'                            : param 为窗口点数（自动调为奇数）
% -------------------------------------------------------------------------
if nargin < 3 || isempty(param)
    param = 0.1;
end

[m,n] = size(M);
S = zeros(m,n);
mth = lower(method);

for k = 1:n
    y = M(:,k);
    switch mth
        case {'loess','lowess','rloess','rlowess'}
            span = param;
            if span <= 0
                span = 0.05;
            elseif span > 1
                span = span / m;
            end
            S(:,k) = smooth(y, span, mth);

        case 'movmean'
            win = param;
            if win < 1
                win = max(3, round(m * win));
            else
                win = round(win);
            end
            S(:,k) = smoothdata(y, 'movmean', win, 'omitnan');

        case 'sgolay'
            win = param;
            if win < 3
                win = 5;
            else
                win = round(win);
            end
            if mod(win,2) == 0
                win = win + 1;
            end
            S(:,k) = smoothdata(y, 'sgolay', win, 'omitnan');

        otherwise
            error('未知平滑方法：%s', method);
    end
end
end

function summary = analyzeSpectra_Batch( ...
    data_x, data_y_raw, data_y_filt, headers, target_dir, base_name, params)
% -------------------------------------------------------------------------
% analyzeSpectra_Batch
% 功能：
%   - 对每一列光谱进行：
%       · 原始 & 平滑光谱叠加绘制
%       · 峰值、FWHM、积分面积计算
%   - 汇总到 summary cell 数组
% -------------------------------------------------------------------------
n_col = size(data_y_raw,2);
summary = cell(n_col+1, 9);
summary(1,:) = {'Power/Condition','X_raw','Y_raw','FWHM_raw', ...
                               'X_filt','Y_filt','FWHM_filt','Area_raw','Area_filt'};

[rows, cols] = subPlotLayout(n_col);

if params.show_figure
    hFig = figure('Name','Per-Column Analysis','WindowState','maximized');
else
    hFig = figure('Visible','off','Name','Per-Column Analysis');
end
clf(hFig);

for i = 1:n_col
    subplot(rows, cols, i); hold on;

    yR = data_y_raw(:,i);
    yF = data_y_filt(:,i);

    % 背景填充（平滑结果）
    area(data_x, yF, 'FaceAlpha',0.18, ...
         'FaceColor',[1 0.4 0.4], 'EdgeColor','none');

    % 原始（蓝）
    plot(data_x, yR, 'b','LineWidth',1.2);
    % 平滑（红虚线）
    plot(data_x, yF, 'r--','LineWidth',1.3);

    % 标题
    titleStr = headers{i};
    if isnumeric(titleStr)
        title(sprintf('%.4g', titleStr));
    else
        title(char(string(titleStr)));
    end

    % 峰值
    [xR,yRmax] = markMax(yR, data_x, 'bo', [-0.2,0.1]);
    [xF,yFmax] = markMax(yF, data_x, 'r^', [-0.2,0.1]);

    % FWHM
    fR = markFWHM(yR, data_x, 'g-', true);
    fF = markFWHM(yF, data_x, 'm--', true);

    % 积分面积
    aR = safeTrapz(data_x, yR);
    aF = safeTrapz(data_x, yF);

    summary(i+1,:) = {headers{i}, xR,yRmax,fR, xF,yFmax,fF, aR,aF};

    xlabel('Energy (eV)');
    ylabel('Intensity (a.u.)');
    grid on;
    hold off;
end

sgtitle('Power / Condition Dependence','FontSize',14);

if params.save_plot
    saveas(hFig, fullfile(target_dir, ['1_SingleColumnPlots_' base_name '.png']));
end
end

function plotSummary_Batch(summary, target_dir, base_name, params)
% -------------------------------------------------------------------------
% plotSummary_Batch
% 功能：根据 summary 绘制：
%   - Peak Energy (filtered) vs Power
%   - Peak Intensity (filtered) vs Power
%   - Area (filtered) vs Power
% -------------------------------------------------------------------------
powerVals = cell2mat(summary(2:end,1));
E_filt    = cell2mat(summary(2:end,5));
I_filt    = cell2mat(summary(2:end,6));
Area_filt = cell2mat(summary(2:end,9));

if params.show_figure
    hFig2 = figure('Name','Summary Plots','WindowState','maximized');
else
    hFig2 = figure('Visible','off','Name','Summary Plots');
end

subplot(1,3,1);
plot(powerVals, E_filt, '-o','LineWidth',1.2);
xlabel('Power / Condition');
ylabel('Peak Energy (eV)');
grid on;

subplot(1,3,2);
plot(powerVals, I_filt, '-s','LineWidth',1.2);
xlabel('Power / Condition');
ylabel('Peak Intensity (a.u.)');
grid on;

subplot(1,3,3);
plot(powerVals, Area_filt, '-^','LineWidth',1.2);
xlabel('Power / Condition');
ylabel('Area (filtered)');
grid on;

if params.save_plot
    saveas(hFig2, fullfile(target_dir, ['2_SummaryPlots_' base_name '.png']));
end
end

function exportPublicationReady_Batch(summary, target_dir)
% -------------------------------------------------------------------------
% exportPublicationReady_Batch
% 功能：
%   - 输出三个 CSV：
%       · PeakEnergy.csv     (Power, PeakEnergy_eV)
%       · PeakIntensity.csv  (Power, PeakIntensity)
%       · Area.csv           (Power, Area)
% -------------------------------------------------------------------------
pubDir = fullfile(target_dir, '6_Publication_Ready');
if exist(pubDir,'dir'); rmdir(pubDir,'s'); end
mkdir(pubDir);

powerVals = cell2mat(summary(2:end,1));
E_filt    = cell2mat(summary(2:end,5));
I_filt    = cell2mat(summary(2:end,6));
Area_filt = cell2mat(summary(2:end,9));

T1 = table(powerVals, E_filt, 'VariableNames', {'Power','PeakEnergy_eV'});
T2 = table(powerVals, I_filt, 'VariableNames', {'Power','PeakIntensity'});
T3 = table(powerVals, Area_filt, 'VariableNames', {'Power','Area'});

writetable(T1, fullfile(pubDir, 'PeakEnergy.csv'));
writetable(T2, fullfile(pubDir, 'PeakIntensity.csv'));
writetable(T3, fullfile(pubDir, 'Area.csv'));

fprintf('\n[Publication-Ready] 数据已生成于：%s\n\n', pubDir);
end

function saveAllOutput_Batch(summary, data_x, data_y_filt, headers, target_dir, base_name, params)
% -------------------------------------------------------------------------
% saveAllOutput_Batch
% 功能：
%   - 保存原始 Data（处理 missing）
%   - 保存滤波光谱（Energy_X + 各列）
%   - 保存 summary 结果
%   - 保存参数日志
% -------------------------------------------------------------------------
% 1) 保存滤波光谱
yFiltOut = [data_x, data_y_filt];
headerFilt = [{'Energy_X'}, headers];
writecell([headerFilt; num2cell(yFiltOut)], ...
    fullfile(target_dir, ['4_Filtered_' base_name '.xlsx']));

% 2) 保存 summary
writecell(summary, fullfile(target_dir, ['5_ResultSummary_' base_name '.xlsx']));

% 3) 参数日志
logFile = fullfile(target_dir, ['0_Parameters_' base_name '.txt']);
fid = fopen(logFile,'w');
dt = datetime('now', 'Format', 'yyyy-MM-dd HH:mm:ss');

fprintf(fid, 'Processing date       : %s\n',  char(dt));
fprintf(fid, 'Data file             : %s\n',  [base_name, ' (Batch Spectra)']);
fprintf(fid, 'Excel row range       : %d → %d\n', params.row_start, params.row_end);
fprintf(fid, 'X axis range          : %.6f → %.6f\n', data_x(1), data_x(end));
fprintf(fid, 'Baseline value        : %.6f\n', params.baseline);
fprintf(fid, 'Smoothing method      : %s\n', params.smooth_type);
fprintf(fid, 'Smoothing parameter   : %.6g\n', params.smooth_param);
fprintf(fid, 'Number of columns     : %d\n', size(data_y_filt,2));
fclose(fid);
end

function [rows, cols] = subPlotLayout(n)
% -------------------------------------------------------------------------
% subPlotLayout
% 功能：给定 n 个子图，返回合理 rows × cols 布局
% -------------------------------------------------------------------------
rows = ceil(sqrt(n));
cols = ceil(n/rows);
end

function [x0,y0] = markMax(y, x, style, offset)
% -------------------------------------------------------------------------
% markMax
% 功能：标注并返回给定谱线的最大值位置
% -------------------------------------------------------------------------
if nargin < 4
    offset = [0,0];
end
if nargin < 3
    style = 'ro';
end

[y0, idx] = max(y);
x0 = x(idx);

plot(x0, y0, style, 'MarkerSize',7, 'LineWidth',1.2);
text(x0 + offset(1), y0 + offset(2), ...
     sprintf('(%.3f, %.3f)', x0, y0), ...
     'VerticalAlignment','bottom', 'HorizontalAlignment','right');
end

function F = markFWHM(y, x, style, showLabel)
% -------------------------------------------------------------------------
% markFWHM
% 功能：标注并返回 FWHM（基于最大值一半）
% -------------------------------------------------------------------------
if nargin < 3
    style = 'g--';
end
if nargin < 4
    showLabel = true;
end

if all(isnan(y)) || max(y) <= 0
    F = NaN;
    return;
end

halfMax = max(y) / 2;
idx = y >= halfMax;
if ~any(idx)
    F = NaN;
    return;
end

xL = x(find(idx,1,'first'));
xR = x(find(idx,1,'last'));
F  = abs(xR - xL);

plot([xL, xR], [halfMax, halfMax], style, 'LineWidth',1.0);
if showLabel
    text(xL, halfMax, sprintf('FWHM = %.3f', F), ...
         'VerticalAlignment','bottom','HorizontalAlignment','left');
end
end

function A = safeTrapz(x, y)
% -------------------------------------------------------------------------
% safeTrapz
% 功能：对谱线进行积分，并确保 x 单调升序与长度匹配
% -------------------------------------------------------------------------
x = x(:);
y = y(:);
L = min(numel(x), numel(y));
x = x(1:L);
y = y(1:L);

if x(1) > x(end)
    x = flip(x);
    y = flip(y);
end

A = abs(trapz(x, y));
end
