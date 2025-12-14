% =========================================================================
% Circular Polarization & g-factor Analysis (Unified Framework v1.1)
%
% 功能：
%   - 从单个数据文件（Excel/CSV/TXT）读取圆偏振光谱数据
%   - 对 Positive / Negative 光谱进行基线校正与平滑
%   - 计算每个磁场下的峰值、FWHM、积分面积（Raw & Smooth）
%   - 计算 DOCP 与 Zeeman splitting，并拟合 g 因子
%   - 绘制各类光谱图、DOCP & Zeeman 图
%   - 导出原始/滤波光谱、summary 表和 Publication-Ready CSV
%
% 约定：
%   - 文件首行包含列标题；第 2 列开始为光谱数据，对应不同磁场 B
%   - 数据区为 Excel 绝对行号 [row_start:row_end]
%   - 第 1 列为能量（或波长）X 轴
%   - 第 2,4,6,... 列为 Positive；3,5,7,... 列为 Negative
%   - 磁场 B 从标题行解析，然后统一按从负到正升序排序
%
% 作者：Shikun Hou
% 版本：Unified Framework v1.1
% 更新时间：2025-11-26
% =========================================================================

clc; clear; close all;
tic;

%% ========================================================================
% 1. 用户参数设置（Parameters）
% ========================================================================
params = struct();

% 注意：以下行号为 Excel 中的"绝对行号"（包含标题行）
params.row_start      = 150;       % 数据起始行（Excel 行号）
params.row_end        = 171;      % 数据终止行（Excel 行号）

params.baseline       = 0;       % 基线值（从所有 Y 列减去）
params.smooth_type    = 'loess';   % 'loess','lowess','movmean','sgolay',...
params.smooth_param   = 0.001;     % 平滑参数（相对跨度或窗口大小，视方法而定）

% 合并 summary 时用于标记 Positive/Negative 的占位值
params.posneg_tag     = [450, 4590];

params.save_plot      = true;      % 是否保存图像
params.show_figure    = true;      % 是否在屏幕显示图像

%% ========================================================================
% 2. 文件选择与输出目录设置（IO）
% ========================================================================
[file_path, target_dir, base_name] = setupIO();

%% ========================================================================
% 3. 数据读取与预处理（Import & Preprocess）
%    包括磁场 B 的解析和从负到正的排序
% ========================================================================
[energy_x, pos_raw, neg_raw, B_fields, raw_header_row1, raw_header_row2] = importDataCP(file_path, params);

%% ========================================================================
% 4. 数据平滑（Smoothing）
% ========================================================================
[pos_filt, neg_filt] = applySmoothingCP(pos_raw, neg_raw, params);

%% ========================================================================
% 5. 单列分析与光谱绘制（Peak, FWHM, Area, Spectra Plots）
% ========================================================================
cp_result = analyzeSpectraCP( ...
    energy_x, pos_raw, neg_raw, pos_filt, neg_filt, B_fields, ...
    target_dir, base_name, params);

%% ========================================================================
% 6. DOCP & g-factor 绘图（Summary Plots）
% ========================================================================
plotSummaryCP(cp_result, target_dir, base_name, params);

%% ========================================================================
% 7. 论文绘图友好数据导出（Publication-Ready Data）
% ========================================================================
exportPublicationReadyCP(cp_result, target_dir);

%% ========================================================================
% 8. 保存原始/滤波数据、summary 与参数日志（Save All Output）
% ========================================================================
saveAllOutputCP(cp_result, energy_x, pos_raw, neg_raw, ...
    pos_filt, neg_filt, B_fields, ...
    raw_header_row1, raw_header_row2, ...
    target_dir, base_name, params);

fprintf('Circular Polarization & g-factor 分析完成，总耗时：%.2f 秒\n', toc);

% =========================================================================
% 9. 本脚本使用的本地函数
% =========================================================================

function [file_path, target_dir, base_name] = setupIO()
% -------------------------------------------------------------------------
% setupIO
% 功能：
%   - 选择数据文件
%   - 在同一目录下创建带日期标签的输出文件夹
% 输出：
%   - file_path : 完整数据文件路径
%   - target_dir: 输出目录
%   - base_name : 文件无扩展名的基本名
% -------------------------------------------------------------------------
[filename, pathname] = uigetfile( ...
    {'*.xls;*.xlsx;*.txt;*.csv','文本/表格文件 (*.xls;*.xlsx;*.txt;*.csv)'; ...
     '*.*','所有文件 (*.*)'}, ...
     '请选择圆偏振光谱数据文件');

if isequal(filename,0)
    error('未选择文件，脚本终止。');
end

file_path = fullfile(pathname, filename);
[~, base_name, ~] = fileparts(filename);

date_str   = char(datetime('now','Format','yyyyMMdd'));
folder_tag = sprintf('%s_%s', date_str, base_name);
target_dir = fullfile(pathname, folder_tag);

if exist(target_dir,'dir')
    rmdir(target_dir,'s');
end
mkdir(target_dir);
end

function [energy_x, pos_raw, neg_raw, B_fields, raw_header_row1, raw_header_row2] = importDataCP(file_path, params)
% -------------------------------------------------------------------------
% importDataCP
% 功能：
%   - 读取圆偏振光谱文件
%   - 从标题行解析磁场 B
%   - 根据 Excel 行号截取数据
%   - 拆分 Positive（偶列）和 Negative（奇列）
%   - 最终将 B 从负到正升序排序，并同步排序所有光谱
% 输出：
%   - energy_x : 能量轴（列向量）
%   - pos_raw  : Positive 原始数据矩阵
%   - neg_raw  : Negative 原始数据矩阵
%   - B_fields : 排序后的磁场数组（从负到正）
% -------------------------------------------------------------------------
Data = readcell(file_path);

raw_header_row1 = Data(1, 2:end);   % 第一行（除掉能量列）
raw_header_row2 = Data(2, 2:end);   % 第二行（完整光谱信息）

n_rows = size(Data,1);
if params.row_start < 2 || params.row_end > n_rows || params.row_start >= params.row_end
    error('行号设置不合法：row_start=%d, row_end=%d, 总行数=%d', ...
        params.row_start, params.row_end, n_rows);
end

% 假定第一行是标题行：其中第 2,3,4,... 列包含与 B 相关信息
title_block = Data(1, :);
header_row  = title_block(2:end);  % 从第二列开始解析 B

% 从标题中解析磁场，并保持长度与 header_row 一致
B_raw = parseMagneticFieldFull(header_row);   % 这里不丢弃 NaN，占位保持

% 截取数据区
data_block = Data(params.row_start:params.row_end, :);
energy_x   = cell2mat(data_block(:, 1));
data_y_all = cell2mat(data_block(:, 2:end));

% 确保列数一致
n_cols_all = size(data_y_all, 2);
if numel(B_raw) < n_cols_all
    B_raw(numel(B_raw)+1:n_cols_all) = NaN;
elseif numel(B_raw) > n_cols_all
    B_raw = B_raw(1:n_cols_all);
end

% 假定 2,4,6,... 为 Positive；3,5,7,... 为 Negative
idx_pos = 1:2:n_cols_all;
idx_neg = 2:2:n_cols_all;
idx_pos = idx_pos(idx_pos <= n_cols_all);
idx_neg = idx_neg(idx_neg <= n_cols_all);

pos_raw_all = data_y_all(:, idx_pos) - params.baseline;
neg_raw_all = data_y_all(:, idx_neg) - params.baseline;

% 对应磁场（按列匹配）
B_pos = B_raw(idx_pos);
B_neg = B_raw(idx_neg);

% 只保留 B 有效的列（非 NaN）
valid_pos = ~isnan(B_pos);
valid_neg = ~isnan(B_neg);

pos_raw = pos_raw_all(:, valid_pos);
neg_raw = neg_raw_all(:, valid_neg);

B_pos = B_pos(valid_pos);
B_neg = B_neg(valid_neg);

% 对于 CP 分析，要求正负磁场一一对应，取两者共有的 B 值
[B_intersect, ia_pos, ia_neg] = intersect(B_pos, B_neg);

if isempty(B_intersect)
    error('Positive 与 Negative 光谱对应的磁场 B 没有交集，请检查标题行。');
end

% 按交集匹配
pos_raw = pos_raw(:, ia_pos);
neg_raw = neg_raw(:, ia_neg);
B_fields = B_intersect(:).';      % 行向量

% 统一长度 n_series
n_series = min([size(pos_raw,2), size(neg_raw,2), numel(B_fields)]);
pos_raw  = pos_raw(:, 1:n_series);
neg_raw  = neg_raw(:, 1:n_series);
B_fields = B_fields(1:n_series);

% === 关键：对 B 从负到正升序排序，并同步排序所有光谱 ===
[B_sorted, idx_sort] = sort(B_fields, 'ascend');   % 负 → 正
B_fields = B_sorted(:);                            % 列向量
pos_raw  = pos_raw(:, idx_sort);
neg_raw  = neg_raw(:, idx_sort);

energy_x = energy_x(:);
end

function B = parseMagneticFieldFull(header_cells)
% -------------------------------------------------------------------------
% parseMagneticFieldFull
% 功能：
%   - 从标题单元格中解析磁场数值，例如 'B=-9T', '-6T', '1', '1.0T'
%   - 返回与 header_cells 等长的数组，未识别的填 NaN
% -------------------------------------------------------------------------
n = numel(header_cells);
B = nan(1, n);

for k = 1:n
    v = header_cells{k};
    if isnumeric(v)
        B(k) = v;
    elseif ischar(v) || isstring(v)
        token = regexp(char(v), '([-+]?\d+\.?\d*)', 'tokens', 'once');
        if ~isempty(token)
            B(k) = str2double(token{1});
        else
            B(k) = NaN;
        end
    else
        B(k) = NaN;
    end
end
end

function [pos_filt, neg_filt] = applySmoothingCP(pos_raw, neg_raw, params)
% -------------------------------------------------------------------------
% applySmoothingCP
% 功能：对 Positive / Negative 数据矩阵进行平滑
% -------------------------------------------------------------------------
pos_filt = smoothMatrixGeneric(pos_raw, params.smooth_type, params.smooth_param);
neg_filt = smoothMatrixGeneric(neg_raw, params.smooth_type, params.smooth_param);
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
            warning('未知平滑方法：%s，本列不平滑。', method);
            S(:,k) = y;
    end
end
end

function cp_result = analyzeSpectraCP( ...
    energy_x, pos_raw, neg_raw, pos_filt, neg_filt, B_fields, ...
    target_dir, base_name, params)
% -------------------------------------------------------------------------
% analyzeSpectraCP
% 功能：
%   - 对每个 B 场下的 Positive / Negative 光谱进行：
%       · 原始 & 平滑光谱叠加绘制
%       · 峰值、FWHM、积分面积计算
%   - 生成 summaryPos / summaryNeg
%   - 绘制 Pos-only, Neg-only, Pos&Neg 叠加图
%   - 所有内容均基于已排序好的 B_fields（负 → 正）
% -------------------------------------------------------------------------
n_series = numel(B_fields);
X = energy_x(:);

% 初始化 summary 表
summaryPos = cell(n_series+1, 9);
summaryNeg = cell(n_series+1, 9);
summaryPos(1,:) = {'B','E_raw','I_raw','FWHM_raw','E_smooth','I_smooth','FWHM_smooth','Area_raw','Area_smooth'};
summaryNeg(1,:) = summaryPos(1,:);

for i = 1:n_series
    summaryPos{i+1,1} = B_fields(i);
    summaryNeg{i+1,1} = B_fields(i);
end

[rows, cols] = subPlotLayout(n_series);

% ---------- Positive 图 ----------
if params.show_figure
    figPos = figure('Name','Positive Spectra','WindowState','maximized');
else
    figPos = figure('Visible','off','Name','Positive Spectra');
end

for i = 1:n_series
    subplot(rows, cols, i); hold on;

    yR = pos_raw(:,i);
    yF = pos_filt(:,i);

    % 积分面积（只对 >0 部分）
    maskR = yR > 0;
    maskF = yF > 0;
    aR = trapz(X(maskR), yR(maskR));
    aF = trapz(X(maskF), yF(maskF));

    area(X(maskF), yF(maskF), 'FaceColor',[1 0.7 0.7], 'FaceAlpha',0.35, 'EdgeColor','none');
    area(X(maskR), yR(maskR), 'FaceColor',[0.8 0.8 0.8], 'FaceAlpha',0.3,  'EdgeColor','none');

    plot(X, yR, 'Color',[0.2 0.2 0.2], 'LineWidth',1.1);
    plot(X, yF, 'r--', 'LineWidth',1.3);

    % 峰值 & FWHM
    % [xR,yRmax] = markMax(yR, X, 'ko', [-0.2,0.1]);
    % [xF,yFmax] = markMax(yF, X, 'r*', [-0.2,0.1]);
    [xR,yRmax] = peakDetectEnhanced(yR, X, 'ko', [-0.2,0.1]);
    [xF,yFmax] = peakDetectEnhanced(yF, X, 'r*', [-0.2,0.1]);



    fR = markFWHM(yR, X, 'g--', false);
    fF = markFWHM(yF, X, 'b--', false);

    summaryPos(i+1,2:9) = {xR,yRmax,fR, xF,yFmax,fF, aR,aF};

    title(sprintf('Positive   B = %.2f T', B_fields(i)));
    xlabel('Energy (eV)');
    ylabel('Intensity (a.u.)');
    grid on; hold off;
end
sgtitle('Positive Spectra (Raw vs Smoothed)');
if params.save_plot
    saveas(figPos, fullfile(target_dir, ['1_Pos_' base_name '.png']));
end

% ---------- Negative 图 ----------
if params.show_figure
    figNeg = figure('Name','Negative Spectra','WindowState','maximized');
else
    figNeg = figure('Visible','off','Name','Negative Spectra');
end

for i = 1:n_series
    subplot(rows, cols, i); hold on;

    yR = neg_raw(:,i);
    yF = neg_filt(:,i);

    maskR = yR > 0;
    maskF = yF > 0;
    aR = trapz(X(maskR), yR(maskR));
    aF = trapz(X(maskF), yF(maskF));

    area(X(maskF), yF(maskF), 'FaceColor',[0.8 0.8 1], 'FaceAlpha',0.35, 'EdgeColor','none');
    area(X(maskR), yR(maskR), 'FaceColor',[0.6 0.6 1], 'FaceAlpha',0.3,  'EdgeColor','none');

    plot(X, yR, 'b',  'LineWidth',1.1);
    plot(X, yF, 'c--','LineWidth',1.3);

    % [xR,yRmax] = markMax(yR, X, 'bo', [-0.2,0.1]);
    % [xF,yFmax] = markMax(yF, X, 'c*', [-0.2,0.1]);


    [xR,yRmax] = peakDetectEnhanced(yR, X, 'bo', [-0.2,0.1]);
    [xF,yFmax] = peakDetectEnhanced(yF, X, 'c*', [-0.2,0.1]);

    fR = markFWHM(yR, X, 'g--', false);
    fF = markFWHM(yF, X, 'b--', false);

    summaryNeg(i+1,2:9) = {xR,yRmax,fR, xF,yFmax,fF, aR,aF};

    title(sprintf('Negative   B = %.2f T', B_fields(i)));
    xlabel('Energy (eV)');
    ylabel('Intensity (a.u.)');
    grid on; hold off;
end
sgtitle('Negative Spectra (Raw vs Smoothed)');
if params.save_plot
    saveas(figNeg, fullfile(target_dir, ['2_Neg_' base_name '.png']));
end

% ---------- Pos & Neg 叠加图（加入峰值标注） ----------
if params.show_figure
    figPN = figure('Name','Pos & Neg Overlap','WindowState','maximized');
else
    figPN = figure('Visible','off','Name','Pos & Neg Overlap');
end

for i = 1:n_series
    subplot(rows, cols, i); hold on;

    yNegR = neg_raw(:,i);
    yNegF = neg_filt(:,i);
    yPosR = pos_raw(:,i);
    yPosF = pos_filt(:,i);

    % 画曲线
    plot(X, yNegR,'b','LineWidth',1.1);
    plot(X, yNegF,'c--','LineWidth',1.3);
    plot(X, yPosR,'k','LineWidth',1.1);
    plot(X, yPosF,'r--','LineWidth',1.3);

    % === 增加峰值符号标注（与 Fig1/2 保持一致） ===
    % [~,~] = markMax(yNegR, X, 'bo', [-0.2,0.1]);
    % [~,~] = markMax(yNegF, X, 'c*', [-0.2,0.1]);
    % [~,~] = markMax(yPosR, X, 'ko', [-0.2,0.1]);
    % [~,~] = markMax(yPosF, X, 'r*', [-0.2,0.1]);

    [~,~] = peakDetectEnhanced(yNegR, X, 'bo', [-0.2,0.1]);
    [~,~] = peakDetectEnhanced(yNegF, X, 'c*', [-0.2,0.1]);
    [~,~] = peakDetectEnhanced(yPosR, X, 'ko', [-0.2,0.1]);
    [~,~] = peakDetectEnhanced(yPosF, X, 'r*', [-0.2,0.1]);

    title(sprintf('Pos & Neg   B = %.2f T', B_fields(i)));
    xlabel('Energy (eV)');
    ylabel('Intensity (a.u.)');
    grid on; hold off;
end

sgtitle('Pos & Neg Overlap (Raw & Smoothed)');
if params.save_plot
    saveas(figPN, fullfile(target_dir, ['3_PosNeg_' base_name '.png']));
end

if params.save_plot
    saveas(figPN, fullfile(target_dir, ['3_PosNeg_' base_name '.png']));
end

% === 计算 DOCP & Zeeman splitting & g 因子（基于已排序 B） ===
ePos_raw = cell2mat(summaryPos(2:end,2));
eNeg_raw = cell2mat(summaryNeg(2:end,2));
ePos_flt = cell2mat(summaryPos(2:end,5));
eNeg_flt = cell2mat(summaryNeg(2:end,5));

deltaE_raw  = (ePos_raw - eNeg_raw) * 1e3;   % meV
deltaE_filt = (ePos_flt - eNeg_flt) * 1e3;   % meV

iPos_raw = cell2mat(summaryPos(2:end,3));
iNeg_raw = cell2mat(summaryNeg(2:end,3));
iPos_flt = cell2mat(summaryPos(2:end,6));
iNeg_flt = cell2mat(summaryNeg(2:end,6));

epsVal = 1e-12;
docp_raw  = (iPos_raw - iNeg_raw) ./ max(abs(iPos_raw + iNeg_raw), epsVal);
docp_filt = (iPos_flt - iNeg_flt) ./ max(abs(iPos_flt + iNeg_flt), epsVal);

B = B_fields(:);

% === g-factor with uncertainty (new) ===

% 使用 polyfit 进行一阶拟合（强制过零点的方式：直接拟合 y/x）
slopeFit = sum(B .* deltaE_raw) / sum(B.^2);    % 最优无截距拟合
zeemanFit = slopeFit .* B;

% 计算误差
deltaFit = deltaE_raw - zeemanFit;              % 残差
N = length(B);
slope_err = sqrt( sum(deltaFit.^2) / ((N-1) * sum(B.^2)) );

% g 因子及误差
muB = 0.05788;        % meV/T
g_value    = slopeFit    / muB;
g_error    = slope_err   / muB;

% 显示结果
fprintf('拟合得到的 g 因子：g = %.4f ± %.4f\n', g_value, g_error);

cp_result.B           = B;
cp_result.summaryPos  = summaryPos;
cp_result.summaryNeg  = summaryNeg;
cp_result.deltaE_raw  = deltaE_raw;
cp_result.deltaE_filt = deltaE_filt;
cp_result.zeemanFit   = zeemanFit;
cp_result.docp_raw    = docp_raw;
cp_result.docp_filt   = docp_filt;
cp_result.slope       = slopeFit;
cp_result.g_value     = g_value;
cp_result.g_error     = g_error;       % VERY IMPORTANT
cp_result.slope_error = slope_err;

cp_result.pos_raw     = pos_raw;
cp_result.neg_raw     = neg_raw;
cp_result.pos_filt    = pos_filt;
cp_result.neg_filt    = neg_filt;

end

function plotSummaryCP(cp_result, target_dir, base_name, params)
% -------------------------------------------------------------------------
% plotSummaryCP
% 功能：
%   - 绘制 DOCP vs B
%   - 绘制 Zeeman splitting vs B（含拟合直线）
%   - 均基于已排序的 B（负 → 正）
% -------------------------------------------------------------------------
B           = cp_result.B;
docp_raw    = cp_result.docp_raw;
docp_filt   = cp_result.docp_filt;
deltaE_raw  = cp_result.deltaE_raw;
zeemanFit   = cp_result.zeemanFit;
g_value     = cp_result.g_value;
g_error     = cp_result.g_error;      % ★★ 修复关键点 ★★

if params.show_figure
    fig = figure('Name','DOCP & Zeeman','WindowState','maximized');
else
    fig = figure('Visible','off','Name','DOCP & Zeeman');
end

subplot(2,1,1);
plot(B, docp_raw,'or','MarkerFaceColor','r'); hold on;
plot(B, docp_filt,'-b','LineWidth',1.4);
xlabel('B (T)'); ylabel('DOCP');
legend('raw','smoothed','Location','best');
grid on;

subplot(2,1,2);
plot(B, deltaE_raw,'bo','MarkerFaceColor','b'); hold on;
plot(B, zeemanFit,'r-','LineWidth',1.8);
xlabel('B (T)'); ylabel('\DeltaE (meV)');
legend('Exp','Fit','Location','best');
grid on;

% 标注 g
xpos = min(B) + 0.05*(max(B)-min(B));
ypos = min(deltaE_raw) + 0.80*(max(deltaE_raw)-min(deltaE_raw));
text(xpos, ypos, sprintf('g = %.3f ± %.3f', g_value, g_error), ...
    'FontSize',14, 'FontWeight','bold', ...
    'BackgroundColor',[1 1 1 0.7], 'Margin',4);

sgtitle('DOCP & Zeeman Splitting / g-factor');

if params.save_plot
    saveas(fig, fullfile(target_dir, ['4_DOCP_Zeeman_' base_name '.png']));
end
end


function exportPublicationReadyCP(cp_result, target_dir)
% -------------------------------------------------------------------------
% exportPublicationReadyCP
% 功能：
%   - 导出适合论文绘图的 CSV：
%       · DOCP_vs_B.csv
%       · Zeeman_vs_B.csv
%   - B 已按负 → 正排序
% -------------------------------------------------------------------------
pub_dir = fullfile(target_dir, '6_Publication_Ready');
if exist(pub_dir,'dir'); rmdir(pub_dir,'s'); end
mkdir(pub_dir);

B           = cp_result.B;
docp_raw    = cp_result.docp_raw;
docp_filt   = cp_result.docp_filt;
deltaE_raw  = cp_result.deltaE_raw;
deltaE_filt = cp_result.deltaE_filt;
zeemanFit   = cp_result.zeemanFit;
g_value     = cp_result.g_value;

T_docp = table(B, docp_raw, docp_filt, ...
    'VariableNames', {'B_T','DOCP_raw','DOCP_smooth'});
writetable(T_docp, fullfile(pub_dir,'DOCP_vs_B.csv'));

T_zeeman = table(B, deltaE_raw, deltaE_filt, zeemanFit, ...
    'VariableNames', {'B_T','DeltaE_raw_meV','DeltaE_smooth_meV','DeltaE_fit_meV'});
writetable(T_zeeman, fullfile(pub_dir,'Zeeman_vs_B.csv'));

% 额外写入一个简单 txt 记录 g
fid = fopen(fullfile(pub_dir,'g_factor.txt'),'w');
fprintf(fid, 'g-factor (dimensionless) = %.6f\n', g_value);
fclose(fid);
end

function saveAllOutputCP(cp_result, energy_x, pos_raw, neg_raw, ...
    pos_filt, neg_filt, B_fields, ...
    raw_header_row1, raw_header_row2, ...
    target_dir, base_name, params)
% -------------------------------------------------------------------------
% saveAllOutputCP (Updated v3.0)
%   正确支持输出原始格式（带两行标题）
% -------------------------------------------------------------------------

B = B_fields(:);

% === 保存矩阵原始/滤波 ===
saveMatrixWithHeader(energy_x, pos_raw , B, fullfile(target_dir,['1_Pos_raw_'  base_name '.xlsx']));
saveMatrixWithHeader(energy_x, pos_filt, B, fullfile(target_dir,['2_Pos_filt_' base_name '.xlsx']));
saveMatrixWithHeader(energy_x, neg_raw , B, fullfile(target_dir,['3_Neg_raw_'  base_name '.xlsx']));
saveMatrixWithHeader(energy_x, neg_filt, B, fullfile(target_dir,['4_Neg_filt_' base_name '.xlsx']));

% === Summary ===
writecell(cp_result.summaryPos, fullfile(target_dir,['5_Pos_summary_' base_name '.xlsx']));
writecell(cp_result.summaryNeg, fullfile(target_dir,['6_Neg_summary_' base_name '.xlsx']));

% === 合并 Summary ===
summaryCombined = buildCombinedSummaryCP(cp_result, params.posneg_tag);
writecell(summaryCombined, fullfile(target_dir,['7_Summary_Combined_' base_name '.xlsx']));

% === 导出原始格式（你特别需要的） ===
exportFilteredCroppedLikeOriginal( ...
    energy_x, pos_filt, neg_filt, B_fields, ...
    raw_header_row1, raw_header_row2, ...
    target_dir, base_name);

% === 参数日志 ===
paramFile = fullfile(target_dir, ['0_Parameters_' base_name '.txt']);
fid = fopen(paramFile,'w');
dt   = datetime('now','Format','yyyy-MM-dd HH:mm:ss');
fprintf(fid, 'Processing date       : %s\n',  char(dt));
fprintf(fid, 'Data file             : %s\n',  [base_name, ' (CP & g-factor)']);
fprintf(fid, 'Excel row range       : %d → %d\n', params.row_start, params.row_end);
fprintf(fid, 'X axis range          : %.6f → %.6f\n', energy_x(1), energy_x(end));
fprintf(fid, 'Baseline value        : %.6f\n', params.baseline);
fprintf(fid, 'Smoothing method      : %s\n', params.smooth_type);
fprintf(fid, 'Smoothing parameter   : %.6g\n', params.smooth_param);
fprintf(fid, 'Magnetic fields (T)   : %s\n', mat2str(B',3));
fprintf(fid, 'g-factor (dimensionless) = %.6f\n', cp_result.g_value);
fclose(fid);

end


function saveMatrixWithHeader(X, mat, B, filepath)
% -------------------------------------------------------------------------
% saveMatrixWithHeader
% 功能：保存 Energy + 多列光谱，首行为表头（Energy, B=xxT,...）
%       B 已按负 → 正排序
% -------------------------------------------------------------------------
X = X(:);
[m,~] = size(mat);
if numel(X) ~= m
    L = min(numel(X), m);
    X   = X(1:L);
    mat = mat(1:L,:);
end

B = B(:)';   % 行向量
magLabels = arrayfun(@(b) sprintf('B=%.2fT', b), B, 'UniformOutput', false);
header    = [{'Energy (eV)'}, magLabels];

data = [num2cell(X), num2cell(mat)];
out  = [header; data];

writecell(out, filepath);
end

function combined = buildCombinedSummaryCP(cp_result, posneg_tag)
% -------------------------------------------------------------------------
% buildCombinedSummaryCP
% 功能：
%   - 将 summaryPos / summaryNeg 与 DOCP / Zeeman 信息合并成一个大的 Cell 表
%   - 磁场 B 已经是负 → 正排序
% -------------------------------------------------------------------------
summaryPos = cp_result.summaryPos;
summaryNeg = cp_result.summaryNeg;
docp_raw   = cp_result.docp_raw;
docp_filt  = cp_result.docp_filt;
deltaE_raw = cp_result.deltaE_raw;
deltaE_fl  = cp_result.deltaE_filt;
zeemanFit  = cp_result.zeemanFit;

nP = size(summaryPos,1);
nN = size(summaryNeg,1);
N  = max(nP, nN);

summaryPos(end+1:N,:) = {[]};
summaryNeg(end+1:N,:) = {[]};
gap = repmat({[]}, N, 1);

blockDOCP = [{'DOCP_raw','DOCP_smooth'}; num2cell([docp_raw(:), docp_filt(:)])];
blockDOE  = [{'DeltaE_raw_meV','DeltaE_smooth_meV','DeltaE_fit_meV'}; ...
             num2cell([deltaE_raw(:), deltaE_fl(:), zeemanFit(:)])];

blockDOCP(end+1:N,:) = {[]};
blockDOE(end+1:N,:)  = {[]};

combined = [summaryPos, gap, summaryNeg, gap, blockDOCP, gap, blockDOE];

% 在第一行加入占位 tag（与原脚本兼容）
row0 = repmat({''}, 1, size(combined,2));
row0{1} = posneg_tag(1);
row0{size(summaryPos,2) + 2} = posneg_tag(2);
combined = [row0; combined];
end

function [rows, cols] = subPlotLayout(n)
% -------------------------------------------------------------------------
% subPlotLayout
% 功能：给定 n 个子图，返回合理的 rows × cols 布局
% -------------------------------------------------------------------------
rows = ceil(sqrt(n));
cols = ceil(n/rows);
end

% function [x0,y0] = markMax(y, x, style, offset)
% % -------------------------------------------------------------------------
% % markMax
% % 功能：标注并返回给定谱线的最大值位置
% % -------------------------------------------------------------------------
% if nargin < 4
%     offset = [0,0];
% end
% if nargin < 3
%     style = 'ro';
% end
% 
% [y0, idx] = max(y);
% x0 = x(idx);
% 
% plot(x0, y0, style, 'MarkerSize',7, 'LineWidth',1.2);
% text(x0 + offset(1), y0 + offset(2), ...
%      sprintf('(%.3f, %.3f)', x0, y0), ...
%      'VerticalAlignment','bottom', 'HorizontalAlignment','right');
% end

function F = markFWHM(y, x, style, showLabel)
% -------------------------------------------------------------------------
% markFWHM
% 功能：标注并返回 FWHM（基于最大值一半）
% -------------------------------------------------------------------------
if nargin < 3
    style = 'g--';
end
if nargin < 4
    showLabel = false;
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

function exportFilteredCroppedLikeOriginal( ...
    X, pos_filt, neg_filt, B_fields, ~, raw_header_row2, ...
    target_dir, base_name)

% -------------------------------------------------------------------------
% exportFilteredCroppedLikeOriginal
%
% 功能：
%   完整恢复原始光谱文件的前两行标签（例如 B 行 + PL 信息行）
%   输出裁剪 + 滤波后的光谱数据（两列一组）
%
% 输入：
%   X                ：能量列（裁剪后的）
%   pos_filt         ：平滑 Positive 光谱矩阵 (N × nB)
%   neg_filt         ：平滑 Negative 光谱矩阵 (N × nB)
%   B_fields         ：B 值（升序）
%   raw_header_row1  ：原始文件第 1 行（用于重建前两行）
%   raw_header_row2  ：原始文件第 2 行（完整文字标签）
%
% 输出：
%   <base_name>_Filtered_Reformatted_RawStyle.xlsx
% -------------------------------------------------------------------------

X = X(:);
nRow = length(X);
nB = numel(B_fields);

%% --- 重新构建 Row1（磁场值） ---
row1 = cell(1, 1 + 2*nB);
row1{1} = ''; 
for i = 1:nB
    row1{2*i}   = B_fields(i);
    row1{2*i+1} = B_fields(i);
end

%% --- 重建 Row2（PL…标签） ---
row2 = cell(1, 1 + 2*nB);
row2{1} = '';

% 原始标签要求按 B 排序重新提取
% raw_header_row2: 是 1×(2*nB) cell
for i = 1:2*nB
    row2{i+1} = raw_header_row2{i};
end

%% --- 构建数据区 ---
data = cell(nRow, 1 + 2*nB);
data(:,1) = num2cell(X);

for i = 1:nB
    data(:, 2*i)   = num2cell(pos_filt(:,i)); % Positive
    data(:, 2*i+1) = num2cell(neg_filt(:,i)); % Negative
end

%% --- 合并整个 Excel 表 ---
output = [row1; row2; data];

outfile = fullfile(target_dir, [base_name '_Filtered_Reformatted_RawStyle.xlsx']);
writecell(output, outfile);

fprintf('已生成完全重建原始格式的裁剪+滤波文件：\n  %s\n', outfile);
end


function [x0, y0] = peakDetectEnhanced(y, x, style, offset, varargin)
% -------------------------------------------------------------------------
% peakDetectEnhanced
% 功能：
%   - 增强型寻峰：不仅取最大值，还会寻找局部峰值并验证有效性
%   - 自动过滤噪声、假峰
%   - 支持自定义搜索区间宽度（点数）与峰显著性阈值
%
% 输入：
%   y           : 光谱数据列
%   x           : 横坐标
%   style       : 绘制点样式
%   offset      : 文本偏移量
%   varargin    : 可选参数
%       'LocalWidth', N        寻峰局部范围（默认 8 点）
%       'Threshold',  factor   峰高度需高于局部均值 factor 倍（默认 1.4）
%
% 输出：
%   x0, y0      : 峰位置及峰强度
%
% -------------------------------------------------------------------------

if nargin < 4
    offset = [0,0];
end
if nargin < 3
    style = 'ro';
end

% 默认参数
p = inputParser;
addParameter(p, 'LocalWidth', 8);
addParameter(p, 'Threshold', 1.4);
parse(p, varargin{:});
localWidth = p.Results.LocalWidth;
thFactor   = p.Results.Threshold;

y = y(:);
x = x(:);

N = length(y);

%% ---------- 1. 计算局部极大值（local maxima） ----------
dy = diff(y);
locMaxIdx = find(dy(1:end-1) > 0 & dy(2:end) < 0) + 1;

if isempty(locMaxIdx)
    % 若无局部峰 → fallback 到最大值
    [y0, idx] = max(y);
    x0 = x(idx);
    plot(x0, y0, style, 'MarkerSize',7,'LineWidth',1.2);
    text(x0+offset(1), y0+offset(2), sprintf('(%.3f, %.3f)', x0,y0));
    return;
end

%% ---------- 2. 验证局部峰是否明显高于背景 ----------
validPeaks = [];
peakScores = [];

for k = locMaxIdx'

    % 限定本地窗口
    iL = max(1, k - localWidth);
    iR = min(N, k + localWidth);
    region = y(iL:iR);

    localMean = mean(region);
    localMax  = y(k);

    % 判断峰是否明显大于周围背景
    if localMax >= thFactor * localMean
        validPeaks(end+1) = k; %#ok<AGROW>
        % 峰评分：越高越优先
        peakScores(end+1) = localMax - localMean;
    end
end

%% ---------- 3. 若无有效峰 → fallback ----------
if isempty(validPeaks)
    [y0, idx] = max(y);
    x0 = x(idx);
    plot(x0, y0, style, 'MarkerSize',7,'LineWidth',1.2);
    text(x0+offset(1), y0+offset(2), sprintf('(%.3f, %.3f)', x0,y0));
    return;
end

%% ---------- 4. 选出评分最高的峰 ----------
[~, bestIdx] = max(peakScores);
idx = validPeaks(bestIdx);

y0 = y(idx);
x0 = x(idx);

%% ---------- 5. 绘图 ----------
plot(x0, y0, style, 'MarkerSize',7, 'LineWidth',1.2);
text(x0+offset(1), y0+offset(2), sprintf('(%.3f, %.3f)', x0,y0), ...
    'VerticalAlignment','bottom', 'HorizontalAlignment','right');

end
