% -------------------------------------------------------------------------
%% Circular Polarization & g-factor Analysis Script
% 脚本功能:
%   本脚本用于处理圆偏振光谱数据，包括：
%       1. 数据导入与预处理（基线校正、行截取、列分离）
%       2. 数据平滑（支持 loess, lowess, movmean, sgolay）
%       3. Positive / Negative 光谱绘图，计算 FWHM 与积分面积
%       4. Positive & Negative 光谱叠加对比
%       5. DOCP (Degree of Circular Polarization) 计算
%       6. Zeeman splitting 分析与 g-factor 拟合
%       7. 保存图像、结果表格与参数日志
%
% 输入文件:
%   支持 Excel (*.xls, *.xlsx), CSV (*.csv), 文本文件 (*.txt)
%   第一行需包含列标题，可含磁场信息
%
% 用户参数:
%   startRow       - 数据起始行
%   endRow         - 数据终止行
%   userBaseline   - 基线值
%   smoothType     - 平滑方法
%   smoothParam    - 平滑参数（span 或窗口大小）
%   posNeg         - 合并 summary 时的占位值
%
% 输出结果:
%   1. Positive / Negative 原始及平滑光谱图 (PNG)
%   2. Positive & Negative 叠加图 (PNG)
%   3. DOCP & g-factor 图 (PNG)
%   4. 原始数据与平滑数据表格 (XLSX)
%   5. Positive / Negative summary 表格 (XLSX)
%   6. 合并 summary 表格，含 DOCP / Zeeman / g-factor (XLSX)
%   7. 参数日志 TXT
%
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------


clc; clear; close all;
tic

%% -----------------------------
% 参数设置
%% -----------------------------
startRow       = 600;      % 数据起始行
endRow         = 1200;     % 数据终止行
userBaseline   = 486;      % 用户设定的基线值
smoothType     = 'loess';  % 平滑算法
smoothParam    = 0.2;      % loess span
posNeg         = [450, 4590];   % 合并时占位值

%% -----------------------------
% 文件选择与文件夹准备
%% -----------------------------
[filename, pathname] = uigetfile( ...
    {'*.xls;*.xlsx;*.txt;*.csv','文本/表格 (*.xls;*.xlsx;*.txt;*.csv)'; ...
     '*.*','所有文件 (*.*)'}, ...
    '请选择数据文件');
if isequal(filename,0) || isequal(pathname,0)
    errordlg('未选择文件，脚本终止','错误'); return;
end
Filepath = fullfile(pathname, filename);

%% -----------------------------
% 创建结果目录（带日期）
%% -----------------------------
[~, basename, ~] = fileparts(filename);
dateStr  = char(datetime('now','Format','yyyyMMdd'));    % 20250930
folderTag = [basename '_' dateStr];
targetDir = fullfile(pathname, folderTag);
if exist(targetDir,'dir'); rmdir(targetDir,'s'); end
mkdir(targetDir);

%% -----------------------------
% 读取数据与分列
%% -----------------------------
Data          = readcell(Filepath);
extractTitle = Data(1:2, :);
extractData  = Data(startRow:endRow, :);
energyX      = extractData(:,1);                % 能量 X 轴

% 解析磁场值
tmpY = extractTitle(1,:)';
tmpY = cellfun(@(x) ifelse(isnumeric(x), x, str2double(x)), tmpY, 'UniformOutput',false);
magneticY = cell2mat(tmpY);
magneticY = magneticY(~isnan(magneticY));
[~, idxY]  = unique(magneticY,'stable');
magneticY = magneticY(idxY);

%% -----------------------------
% 列索引与数据矩阵
%% -----------------------------
[~, nColsTotal] = size(extractData);
columnPos = 2:2:nColsTotal;
columnNeg = 3:2:nColsTotal;

availablePosCols = columnPos(columnPos <= nColsTotal);
availableNegCols = columnNeg(columnNeg <= nColsTotal);

matPosRaw = cell2mat(extractData(:, availablePosCols)) - userBaseline;
matNegRaw = cell2mat(extractData(:, availableNegCols)) - userBaseline;

matPosFilt = smoothMatrix(matPosRaw, smoothType, smoothParam);
matNegFilt = smoothMatrix(matNegRaw, smoothType, smoothParam);

matPos = matPosRaw;
matNeg = matNegRaw;

% —— 确定实际系列数
nSeriesCandidates = [ size(matPos,2), size(matNeg,2), numel(magneticY) ];
nSeries = min(nSeriesCandidates);

if nSeries == 0
    errordlg('未找到可用的数据列，请检查输入文件格式与 startRow/endRow 设置。','数据错误');
    return;
end

summaryPos = createSummaryTable(nSeries, magneticY(1:nSeries));
summaryNeg = createSummaryTable(nSeries, magneticY(1:nSeries));

%% -----------------------------
% 绘图：Positive
%% -----------------------------
X = cell2mat(energyX);
[rows, cols] = subPlotLayout(nSeries);
figPos = figure('WindowState','maximized');
for i = 1:nSeries
    subplot(rows,cols,i);
    colRaw  = matPos(:,i);
    colFilt = matPosFilt(:,i);
    % 曲线
    plot(X, colRaw ,'r-'); hold on;
    plot(X, colFilt,'-r');
    % Raw 面积
    mask      = colRaw > 0;
    areaRaw  = trapz(X(mask), colRaw(mask));
    hA1       = area(X(mask), colRaw(mask),'FaceColor',[1 0.7 0.7],'EdgeColor','none');
    hA1.FaceAlpha = .35;
    % Smooth 面积
    mask2     = colFilt > 0;
    areaFilt = trapz(X(mask2), colFilt(mask2));
    hA2       = area(X(mask2), colFilt(mask2),'FaceColor',[0.7 0.7 1],'EdgeColor','none');
    hA2.FaceAlpha = .35;
    % 注释
    [x1,y1] = annotateMaxValue(matPos,      X, i,'b*',[-0.2,0.1]);
    [x2,y2] = annotateMaxValue(matPosFilt, X, i,'b*',[-0.2,0.1]);
    f1 = annotateFWHM(colRaw , X,'g--');
    f2 = annotateFWHM(colFilt, X,'b--');
    % 写 summary
    summaryPos{i+1,2} = x1;   summaryPos{i+1,3} = y1;
    summaryPos{i+1,4} = f1;   summaryPos{i+1,5} = x2;
    summaryPos{i+1,6} = y2;   summaryPos{i+1,7} = f2;
    summaryPos{i+1,8} = areaRaw;
    summaryPos{i+1,9} = areaFilt;
    hold off;
end
sgtitle('Positive Data','FontSize',14);
set(figPos,'MenuBar','none','ToolBar','none');
saveas(figPos, fullfile(targetDir,['0_Plots_Pos_' basename '.png']));

%% -----------------------------
%% 绘图：Negative
%% -----------------------------
figNeg = figure('WindowState','maximized');
for i = 1:nSeries
    subplot(rows, cols, i);
    colRaw  = matNeg(:, i);
    colFilt = matNegFilt(:, i);
    plot(X, colRaw,  'b-'); hold on;
    plot(X, colFilt, '-b');
    maskRaw   = colRaw  > 0;
    areaRaw   = trapz(X(maskRaw),  colRaw(maskRaw));
    hA1 = area(X(maskRaw), colRaw(maskRaw), 'FaceColor', [0.60 0.60 1.00], 'EdgeColor', 'none');
    hA1.FaceAlpha = 0.35;
    maskFilt  = colFilt > 0;
    areaFilt  = trapz(X(maskFilt), colFilt(maskFilt));
    hA2 = area(X(maskFilt), colFilt(maskFilt), 'FaceColor', [0.80 0.80 1.00], 'EdgeColor', 'none');
    hA2.FaceAlpha = 0.35;
    title(sprintf('B = %.2f T', magneticY(i)));
    [x1, y1] = annotateMaxValue(matNeg     , X, i, 'b*', [-0.2, 0.1]);
    [x2, y2] = annotateMaxValue(matNegFilt, X, i, 'b*', [-0.2, 0.1]);
    f1 = annotateFWHM(colRaw ,  X, 'g--');
    f2 = annotateFWHM(colFilt, X, 'b--');
    summaryNeg{i+1, 2} = x1; summaryNeg{i+1, 3} = y1;
    summaryNeg{i+1, 4} = f1; summaryNeg{i+1, 5} = x2;
    summaryNeg{i+1, 6} = y2; summaryNeg{i+1, 7} = f2;
    summaryNeg{i+1, 8} = areaRaw; summaryNeg{i+1, 9} = areaFilt;
    hold off;
end
sgtitle('Negative Data','FontSize',14);
set(figNeg,'MenuBar','none','ToolBar','none');
saveas(figNeg, fullfile(targetDir,['0_Plots_Neg_' basename '.png']));

%% -----------------------------
%% 绘图：Pos & Neg 叠加
%% -----------------------------
figPN = figure('WindowState','maximized');
for i=1:nSeries
    subplot(rows,cols,i);
    plot(X,matNeg(:,i),'-'); hold on;
    plot(X,matPos(:,i),'-');
    plot(X,matPosFilt(:,i),'-b');
    plot(X,matNegFilt(:,i),'-r');
    title(sprintf('B = %.2f T',magneticY(i)));
    hold off;
end
sgtitle('Pos & Neg Overlap','FontSize',14);
set(figPN,'MenuBar','none','ToolBar','none');
saveas(figPN, fullfile(targetDir,['0_Plots_PosNeg_' basename '.png']));

%% -----------------------------
%% 计算 DOCP 与 g 因子 & 绘图
%% -----------------------------
eRawPos   = cell2mat(summaryPos(2:end,2));
eRawNeg   = cell2mat(summaryNeg(2:end,2));
deltaERaw = eRawPos  - eRawNeg;
eFiltPos  = cell2mat(summaryPos(2:end,5));
eFiltNeg  = cell2mat(summaryNeg(2:end,5));
deltaEFilt= eFiltPos - eFiltNeg;

iRawPos = cell2mat(summaryPos(2:end,3));
iRawNeg = cell2mat(summaryNeg(2:end,3));
docpRaw  = (iRawPos - iRawNeg) ./ (iRawPos + iRawNeg);
iFiltPos= cell2mat(summaryPos(2:end,6));
iFiltNeg= cell2mat(summaryNeg(2:end,6));
docpFilt = (iFiltPos - iFiltNeg) ./ (iFiltPos + iFiltNeg);

zeemanShifts = deltaERaw * 1e3;   % meV
B             = magneticY(1:nSeries);
alpha         = 17.2759858;
fitfun        = @(g,B) g.*B;
g0            = 2;
opts          = optimoptions('lsqcurvefit','Display','off');
gFit         = lsqcurvefit(fitfun, g0, B, zeemanShifts, [], [], opts);
zeemanFit    = fitfun(gFit, B);
Slope         = alpha * gFit;
fprintf('拟合得到的 g 因子: %.4f\n', Slope);

figDocp = figure('WindowState','maximized');
subplot(2,1,1);
plot(B,docpRaw,'or','MarkerFaceColor','r'); hold on;
plot(B,docpFilt,'-b');
xlabel('Magnetic Field (T)'); ylabel('DOCP');
legend('raw','filt','Location','best'); grid on;
subplot(2,1,2);
plot(B,zeemanShifts,'bo','MarkerFaceColor','b'); hold on;
plot(B,zeemanFit,'r-','LineWidth',2);
xlabel('Magnetic Field (T)'); ylabel('\DeltaE (meV)');
legend('Exp','Fit','Location','best'); grid on;
text(max(B)*0.3, max(zeemanShifts)*0.4, sprintf('g = %.2f',Slope), ...
     'FontSize',12,'Color','r','FontWeight','bold');
sgtitle('DOCP & g-factor','FontSize',14);
set(figDocp,'MenuBar','none','ToolBar','none');
saveas(figDocp, fullfile(targetDir,['0_Plots_DOCP_' basename '.png']));

%% -----------------------------
%% 保存数据与结果
%% -----------------------------
saveMatrixWithHeader(energyX, matPos,      magneticY(1:nSeries), fullfile(targetDir,['1_Pos_raw_' basename '.xlsx']));
saveMatrixWithHeader(energyX, matPosFilt, magneticY(1:nSeries), fullfile(targetDir,['2_Pos_filt_' basename '.xlsx']));
writecell(summaryPos, fullfile(targetDir,['3_Pos_summary_' basename '.xlsx']));
saveMatrixWithHeader(energyX, matNeg,      magneticY(1:nSeries), fullfile(targetDir,['4_Neg_raw_' basename '.xlsx']));
saveMatrixWithHeader(energyX, matNegFilt, magneticY(1:nSeries), fullfile(targetDir,['5_Neg_filt_' basename '.xlsx']));
writecell(summaryNeg, fullfile(targetDir,['6_Neg_summary_' basename '.xlsx']));

%% -----------------------------
%% 合并 summaryPos & summaryNeg 并追加 DOCP / Zeeman
%% -----------------------------
nRowPos = size(summaryPos, 1);
nRowNeg = size(summaryNeg, 1);
nRowMax = max(nRowPos, nRowNeg);

if nRowPos < nRowMax, summaryPos(end+1:nRowMax, :) = {[]}; end
if nRowNeg < nRowMax, summaryNeg(end+1:nRowMax, :) = {[]}; end

gapCol     = repmat({[]}, nRowMax, 1);
baseTable  = [summaryPos, gapCol, summaryNeg];

headerDocp   = {'DOCP_raw','DOCP_filt'};
headerZeeman = {'\DeltaE_raw (meV)','\DeltaE_smooth (meV)','\DeltaE_fit (meV)'};

blockDocp   = [headerDocp ; num2cell(docpRaw) , num2cell(docpFilt)];
blockZeeman = [headerZeeman ;
                num2cell(zeemanShifts) , ...
                num2cell(deltaEFilt*1e3) , ...
                num2cell(zeemanFit) ];

blkRows = nRowMax - size(blockDocp,1);
if blkRows > 0, blockDocp(end+1:nRowMax,:)   = {[]}; end
blkRows = nRowMax - size(blockZeeman,1);
if blkRows > 0, blockZeeman(end+1:nRowMax,:) = {[]}; end

summaryResult = [ baseTable, gapCol, blockDocp, gapCol, blockZeeman ];

placeRow = repmat({''}, 1, size(summaryResult,2));
placeRow{1} = posNeg(1);
placeRow{size(summaryPos,2)+2} = posNeg(2);
summaryResult = [placeRow ; summaryResult];

outpathCombined = fullfile(targetDir, ['7_Summary_Combined_' basename '.xlsx']);
writecell(summaryResult, outpathCombined);

%% -----------------------------
%% 保存运行参数到 TXT
%% -----------------------------
xStart = energyX{1};
xEnd   = energyX{end};

paramFile = fullfile(targetDir, ['Parameters_' basename '.txt']);
fid = fopen(paramFile, 'w');
runTime = toc;
dtNow   = datetime('now','Format','yyyy-MM-dd HH:mm:ss');

fprintf(fid, "Date           : %s\n", char(dtNow));
fprintf(fid, "File           : %s\n", filename);
fprintf(fid, "Rows           : %d – %d\n", startRow, endRow);
fprintf(fid, "X-axis range   : Start_X = %.6g, End_X = %.6g\n", xStart, xEnd);
fprintf(fid, "Baseline       : %g\n", userBaseline);
fprintf(fid, "Smooth         : %s (span = %g)\n", smoothType, smoothParam);
fprintf(fid, "Pos_Neg tag    : [%g, %g]\n", posNeg);
fprintf(fid, "Magnetic-Field : %s (T)\n\n", mat2str(magneticY(1:nSeries)',4));

fprintf(fid, "★  全局 g-factor 拟合结果\n");
fprintf(fid, "    g (dimensionless)        : %.6f\n", gFit);
fprintf(fid, "    g·μB/ħ (= Slope, meV/T)  : %.6f\n\n", Slope);

fprintf(fid, "★  额外统计量\n");
fprintf(fid, "    Mean DOCP (raw / filt) : %.5f  /  %.5f\n", ...
        mean(docpRaw,'omitnan'), mean(docpFilt,'omitnan'));
fprintf(fid, "    Script runtime (s)     : %.2f\n", runTime);
fclose(fid);

%% ========================================================================
%% 本脚本所用本地函数
%% ========================================================================
function [rows,cols] = subPlotLayout(n)
    rows = ceil(sqrt(n));
    cols = ceil(n/rows);
end

function out = ifelse(cond, a, b)
    if cond
        out = a;
    else
        out = b;
    end
end

function tbl = createSummaryTable(nCols, Y)
    if numel(Y) < nCols
        Y = [Y(:); nan(nCols - numel(Y), 1)];
    end
    tbl = cell(nCols + 1, 9);
    tbl(1,:) = { ...
        'B (T)', 'E_raw','I_raw','FWHM_raw', ...
        'E_filt','I_filt','FWHM_filt','Area_raw','Area_filt'};
    tbl(2:end,1) = num2cell(Y(1:nCols));
end

function [x,y] = annotateMaxValue(mat,xvals,idx,style,offset)
    if nargin<5, offset=[0,0]; end
    if nargin<4, style='ro'; end
    col = mat(:,idx);
    [y,k] = max(col);
    x = xvals(k);
    plot(x,y,style);
    text(x+offset(1), y+offset(2), sprintf('(%.2f, %.2f)',x,y), 'FontSize',8,'Color','k');
end

function fwhm = annotateFWHM(col, xvals, lineSpec)
    halfMax = max(col) / 2;
    idx = find(col >= halfMax);
    if isempty(idx)
        fwhm = NaN;
        return;
    end
    fwhm = xvals(idx(end)) - xvals(idx(1));
    if nargin > 2
        plot([xvals(idx(1)), xvals(idx(end))],[halfMax, halfMax],lineSpec);
    end
end

function matFilt = smoothMatrix(matRaw, method, span)
    matFilt = zeros(size(matRaw));
    for i=1:size(matRaw,2)
        matFilt(:,i) = smooth(matRaw(:,i), method, span);
    end
end

function saveMatrixWithHeader(X, mat, B, filepath)
    % 保证 X 是列向量
    if iscell(X)
        X = cell2mat(X);
    end
    X = X(:);

    % 保证行数一致
    if size(mat,1) ~= numel(X)
        minLen = min(size(mat,1), numel(X));
        X   = X(1:minLen);
        mat = mat(1:minLen,:);
    end

    % 构建表头
    % 确保B的长度与mat的列数匹配
    nCols = size(mat, 2);
    B = B(:)';  % 强制转换为行向量
    
    % 如果B的长度与mat的列数不匹配，进行调整
    if length(B) > nCols
        B = B(1:nCols);  % 截断多余的B值
    elseif length(B) < nCols
        % 如果B不够，用最后一个值填充
        B = [B, repmat(B(end), 1, nCols - length(B))];
    end
    
    magneticLabels = arrayfun(@(b) sprintf('B=%.2fT', b), B, 'UniformOutput', false);
    header = [{'Energy (eV)'}, magneticLabels];

    % 构建数据区
    data = [num2cell(X), num2cell(mat)];

    % 合并（先 header 行，再数据区）
    out = [header; data];

    % 写出
    writecell(out, filepath);
end

