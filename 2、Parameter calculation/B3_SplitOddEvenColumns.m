%% =========================================================================
%   SplitOddEvenColumns v2.0 (Unified Framework)
%
%   功能：
%       1. 支持 TXT / CSV / XLSX / XLS 文件
%       2. 自动尝试读取表头；如失败则按纯数值矩阵处理
%       3. 自动拆分奇数列与偶数列
%       4. 若存在表头，输出带表头；否则输出纯数字矩阵
%       5. 自动生成 *_odd.xlsx 和 *_even.xlsx
%
%   作者：Shikun Hou
%   更新时间：2025-11-27
%% =========================================================================

clc; clear; close all;

%% ============================ 1. 文件选择 ===============================
[filename, pathname] = uigetfile( ...
    {'*.txt;*.csv;*.xlsx;*.xls','数据文件 (*.txt, *.csv, *.xlsx, *.xls)'}, ...
    '请选择包含至少 2 列数据的文件');

if isequal(filename,0)
    disp('未选择文件，程序终止。');
    return;
end

filepath = fullfile(pathname, filename);

fprintf('\n======================================================\n');
fprintf('      SplitOddEvenColumns v2.0 正在处理文件\n');
fprintf('      文件：%s\n', filename);
fprintf('======================================================\n\n');

%% ============================ 2. 数据读取 ===============================
try
    % 先尝试按 table 读取
    T = readtable(filepath, 'PreserveVariableNames', true);
    data = T{:,:};
    colNames = T.Properties.VariableNames;
    hasHeader = true;

    fprintf('读取方式：table (含表头)\n');

catch
    % 若失败则按矩阵读取
    data = readmatrix(filepath);
    colNames = {};
    hasHeader = false;

    fprintf('读取方式：matrix (无表头)\n');
end

[numRows, numCols] = size(data);

if numCols < 2
    error('数据列数不足：需要至少 2 列才能拆分奇偶列。');
end

fprintf('数据规模：%d 行 × %d 列\n\n', numRows, numCols);

%% ============================ 3. 奇偶列拆分 =============================
oddData  = data(:, 1:2:numCols);      % 奇数列
evenData = data(:, 2:2:numCols);      % 偶数列

if hasHeader
    oddNames  = colNames(1:2:numCols);
    evenNames = colNames(2:2:numCols);

    % 自动修复重复列名（Excel 不允许重复变量名）
    oddNames  = matlab.lang.makeUniqueStrings(oddNames);
    evenNames = matlab.lang.makeUniqueStrings(evenNames);

    T_odd  = array2table(oddData,  'VariableNames', oddNames);
    T_even = array2table(evenData, 'VariableNames', evenNames);
end

%% ============================ 4. 输出文件名 =============================
[~, baseName, ~] = fileparts(filename);

outputOdd  = fullfile(pathname, [baseName '_odd.xlsx']);
outputEven = fullfile(pathname, [baseName '_even.xlsx']);

%% ============================ 5. 写入 Excel =============================
try
    if hasHeader
        writetable(T_odd,  outputOdd);
        writetable(T_even, outputEven);
    else
        writematrix(oddData,  outputOdd);
        writematrix(evenData, outputEven);
    end
catch ME
    error('写入 Excel 文件失败：%s', ME.message);
end

%% ============================ 6. 完成提示 ==============================
fprintf('===================== 拆分完成 =====================\n');
fprintf('奇数列已输出到：%s\n', outputOdd);
fprintf('偶数列已输出到：%s\n', outputEven);
fprintf('=====================================================\n\n');
