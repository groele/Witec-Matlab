% =========================================================================
% FlexiExcelMerge (Unified Framework v1.0)
%
% 功能：
%   支持多文件批量重命名、名称解析、读取 Excel/CSV 数据并自动合并。
%   可用于功率依赖、温度依赖、磁场依赖、电压依赖等多种实验数据整理。
%
% 模块功能结构：
%   (1) 文件名处理：按规则截取、清洗、格式化、去重
%   (2) 文件读取：CSV/XLS/XLSX 自动读取 + 自动补齐行数
%   (3) 数据合并：仅取每文件第一列，按列名排序
%   (4) 输出：MergedData.xlsx
%
% 作者：Shikun Hou
% 版本：Unified Framework v1.0
% 更新时间：2025-11-26
% =========================================================================

clc; clear; close all;
tic;

%% ========================================================================
% 1. 参数设置（Parameters）
%% ========================================================================
params = struct();
params.nameSliceRules = {'end-4:end'};        % 文件名截取规则
params.regexPattern   = '[^0-9.+-]';          % 清除非数字字符
params.sortOption     = 'forward';            % 'forward'/升序，'reverse'/降序

% 常见规则示例（注释说明）
% params.nameSliceRules = {'1:3'};                 % 前3字符
% params.nameSliceRules = {'end-3:end'};           % 倒数4字符
% params.nameSliceRules = {'1:2,end-1:end'};       % 前2 + 后2
% params.nameSliceRules = {'4:6'};                 % 第4到6字符

%% ========================================================================
% 2. 输入文件夹选择（IO）
%% ========================================================================
folderPath = uigetdir(pwd, '请选择数据文件所在文件夹');
if isequal(folderPath, 0)
    error('未选择文件夹，程序终止。');
end
fprintf('已选择文件夹：%s\n', folderPath);

%% ========================================================================
% 3. 批量文件重命名（Filename Processing）
%% ========================================================================
renameFilesInFolder(folderPath, params);

%% ========================================================================
% 4. 删除旧的合并文件（避免冲突）
%% ========================================================================
outputFile = fullfile(folderPath,'MergedData.xlsx');
if exist(outputFile, 'file')
    delete(outputFile);
    fprintf('已删除旧 MergedData.xlsx\n');
end

%% ========================================================================
% 5. 获取所有 Excel/CSV 文件
%% ========================================================================
fileList = collectDataFiles(folderPath);
if isempty(fileList)
    error('文件夹中未找到 Excel/CSV 文件。');
end

%% ========================================================================
% 6. 从所有文件读取数据并准备合并
%% ========================================================================
[validNames, paddedData, maxRows] = preprocessFiles(fileList, folderPath, params);

%% ========================================================================
% 7. 合并所有数据列（首列为 0）
%% ========================================================================
mergedData = mergeColumns(validNames, paddedData, maxRows, params);

%% ========================================================================
% 8. 保存输出
%% ========================================================================
saveMergedResult(outputFile, mergedData);

fprintf('\n全部处理完成！总耗时：%.2f 秒。\n\n', toc);

% =========================================================================
% ========================== 本文件所有本地函数 ===========================
% =========================================================================

%% -------------------------------------------------------------------------
function renameFilesInFolder(folderPath, params)
% renameFilesInFolder
% 功能：
%   批量重命名文件，使用用户自定义截取规则 + 正则清洗
% -------------------------------------------------------------------------
fileList = dir(fullfile(folderPath,'*.*'));
fileList = fileList(~[fileList.isdir]);

if isempty(fileList)
    fprintf('文件夹为空，无需重命名。\n');
    return;
end

fprintf('\n=== 开始批量重命名文件 ===\n');
for k = 1:length(fileList)
    oldFull = fullfile(folderPath, fileList(k).name);
    [~, base, ext] = fileparts(fileList(k).name);

    % 1) 按规则截取文件名
    newName = extractNameByRules(base, params.nameSliceRules);

    % 2) 清理非法字符
    newName = regexprep(newName, params.regexPattern, '');

    if isempty(newName)
        warning('重命名后为空，跳过：%s', fileList(k).name);
        continue;
    end

    newFull = fullfile(folderPath, [newName, ext]);

    % 3) 若目标文件已存在则跳过
    if ~strcmp(oldFull,newFull) && exist(newFull,'file')
        warning('目标文件已存在，跳过：%s', newFull);
        continue;
    end

    try
        movefile(oldFull, newFull);
        fprintf('重命名：%s → %s\n', fileList(k).name, [newName, ext]);
    catch ME
        warning('重命名失败：%s，错误：%s', fileList(k).name, ME.message);
    end
end
fprintf('文件重命名完成。\n');
end

%% -------------------------------------------------------------------------
function fileList = collectDataFiles(folderPath)
% collectDataFiles
% 功能：收集 .csv .xls .xlsx 数据文件
% -------------------------------------------------------------------------
fileList = [
    dir(fullfile(folderPath,'*.csv'));
    dir(fullfile(folderPath,'*.xls'));
    dir(fullfile(folderPath,'*.xlsx'))
];
end

%% -------------------------------------------------------------------------
function [validNames, paddedData, maxRows] = preprocessFiles(fileList, folderPath, params)
% preprocessFiles
% 功能：
%   - 读取所有文件
%   - 自动截取文件名 → 清洗 → 去重
%   - 自动对齐行数（padding NaN）
%   - 多列文件只取第一列
% 输出：
%   validNames  : 各数据列的列名
%   paddedData  : 补齐后的数据列 cell
%   maxRows     : 需要补齐到的最大行数
% -------------------------------------------------------------------------
validNames = {};
paddedData = {};
nameCount  = containers.Map('KeyType','char','ValueType','int32');
maxRows = 0;

for k = 1:length(fileList)
    fileName = fileList(k).name;
    filePath = fullfile(folderPath, fileName);
    [~, base, ~] = fileparts(fileName);

    % === 处理列名 ===
    nameClean = regexprep(base, params.regexPattern, '');
    if isempty(nameClean)
        nameClean = base;
    end
    if strcmp(nameClean,'0')
        nameClean = '0.0';
    end

    % 去重：若重复则加后缀
    if isKey(nameCount, nameClean)
        nameClean = [nameClean, '.1111'];
    end
    nameCount(nameClean) = 1;
    validNames{end+1} = nameClean;

    % === 读取数据 ===
    try
        data = readmatrix(filePath);
        maxRows = max(maxRows, size(data,1));
        paddedData{end+1} = data;
    catch ME
        warning('无法读取文件：%s，错误：%s', fileName, ME.message);
    end
end
end

%% -------------------------------------------------------------------------
function mergedData = mergeColumns(validNames, paddedData, maxRows, params)
% mergeColumns
% 功能：
%   - 所有文件数据按最大行数补齐
%   - 多列文件取第一列
%   - 合并为一个 table
%   - 插入首列 "0" 作为统一格式
%   - 按数值列名排序
% -------------------------------------------------------------------------
mergedData = table();

for k = 1:length(paddedData)
    data = paddedData{k};
    nPad = maxRows - size(data,1);
    if nPad > 0
        data = [data; nan(nPad, size(data,2))];
    end

    if size(data,2) > 1
        data = data(:,1);   % 只取第一列
    end

    T = array2table(data, 'VariableNames', {validNames{k}});
    mergedData = [mergedData, T];
end

% 插入首列 '0'
mergedData = addvars(mergedData, nan(maxRows,1), 'Before',1, 'NewVariableNames', {'0'});

% 排序
varNames = mergedData.Properties.VariableNames;
numericNames = varNames(2:end);
numericVals  = str2double(numericNames);

if strcmp(params.sortOption, 'reverse')
    order = 'descend';
else
    order = 'ascend';
end

[~, idx] = sort(numericVals, order);
sortedNames = [{'0'}, numericNames(idx)];
mergedData = mergedData(:, sortedNames);
end

%% -------------------------------------------------------------------------
function saveMergedResult(outputFile, mergedData)
% saveMergedResult
% 功能：保存合并结果为 Excel
% -------------------------------------------------------------------------
try
    writetable(mergedData, outputFile);
    fprintf('合并完成，已保存：%s\n', outputFile);
catch ME
    error('保存文件失败：%s', ME.message);
end
end

%% -------------------------------------------------------------------------
function newName = extractNameByRules(base, rules)
% extractNameByRules
% 功能：根据规则（如 'end-4:end'）截取文件名片段
% -------------------------------------------------------------------------
newName = '';
n = length(base);

for i = 1:length(rules)
    rule = rules{i};
    rule = strrep(rule, 'end', num2str(n));

    try
        idx = eval(['[', rule, ']']);
    catch
        warning('无效规则 "%s"，跳过', rule);
        continue;
    end

    idx(idx<1 | idx>n) = [];
    newName = [newName, base(idx)];
end
end
