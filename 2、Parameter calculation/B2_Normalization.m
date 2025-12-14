%% =========================================================================
%   Unified Column Normalization Tool  v2.0
%
%   功能：
%       1. 支持 .xlsx / .xls / .csv / .txt 任意格式
%       2. 自动按列归一化（从第 2 列开始）
%       3. 第一行、第一列保持不变（通常是 X 轴或标题）
%       4. 自动处理异常情况（空列、常数列、无法归一化列）
%       5. 输出归一化后的 Excel 文件 *_normalized.xlsx
%
%   作者：Shikun Hou
%   框架：Unified Framework v2.0
%   更新时间：2025-11-27
%% =========================================================================

clc; clear; close all;

%% ======================== 1. 文件选择 =========================
[filename, pathname] = uigetfile( ...
    {'*.xlsx;*.xls;*.csv;*.txt', '数据文件 (*.xlsx, *.xls, *.csv, *.txt)'}, ...
    '请选择数据文件');

if isequal(filename,0)
    disp('未选择文件，程序终止。');
    return;
end

filepath = fullfile(pathname, filename);
[~, baseName, ~] = fileparts(filename);

fprintf('\n===========================================================\n');
fprintf('     Unified Column Normalization Tool v2.0\n');
fprintf('     正在处理文件：%s\n', filename);
fprintf('===========================================================\n\n');

%% ======================== 2. 读取数据 =========================
T = readtable(filepath, 'PreserveVariableNames', true);
data = T{:,:};      % 原始数值矩阵
[numRows, numCols] = size(data);

if numCols < 2
    error('列数不足：至少需要 2 列数据才能归一化。');
end

%% ====================== 3. 按列执行归一化 ======================
normData = data;    % 初始化输出

for col = 2:numCols

    colVec = data(2:end, col);      % 跳过第一行

    % 跳过非数值列（如标题行包含字符串）
    if ~isnumeric(colVec)
        fprintf('跳过第 %d 列（存在非数值数据）。\n', col);
        continue;
    end

    % 排除 NaN
    validVals = colVec(~isnan(colVec));

    if isempty(validVals)
        fprintf('第 %d 列为空（或全为 NaN），设置为 0。\n', col);
        normData(2:end, col) = 0;
        continue;
    end

    minVal = min(validVals);
    maxVal = max(validVals);

    if maxVal == minVal
        fprintf('第 %d 列为常数列，归一化后设为 0。\n', col);
        normData(2:end, col) = 0;
    else
        normData(2:end, col) = (colVec - minVal) ./ (maxVal - minVal);
    end
end

%% ======================== 4. 写回 Excel =========================
T_out = T;
T_out{:,:} = normData;

outputFile = fullfile(pathname, [baseName '_normalized.xlsx']);
writetable(T_out, outputFile);

%% =========================== 5. 完成 =============================
fprintf('\n================== 处理完成 ==================\n');
fprintf('归一化文件已生成：\n  %s\n', outputFile);
fprintf('================================================\n');
