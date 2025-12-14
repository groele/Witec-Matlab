% =========================================================================
% PolarDataManager (Unified Framework v1.1)
%
% 功能：
%   1. 根据 DataType 自动识别电压(V) 或磁场(T)
%   2. 将 CSV 文件移动到以数值命名的子文件夹
%   3. 为每个子文件夹生成汇总 Excel（横向列扩展）
%   4. 生成所有子文件夹的总汇总 Total.xlsx（横向拼接）
%   5. 输出日志文件 ProcessLog.txt
%
% 修复内容（v1.3）：
%   - 修复 combine = cell2mat(dataCells') 导致输出只有一列问题
%   - 修复 Total.xlsx 中列数无法累积的问题
%   - 完全恢复 v13.0 正确的列方向逻辑
%
% 作者：Shikun Hou
% 版本：Unified Framework v1.3
% 更新时间：2025-11-26
% =========================================================================

clc; clear; close all;
tic;

%% ========================================================================
% 1. 参数设置（Parameters）
%% ========================================================================
params = struct();
params.DataType = 'Magnetic';     % 'Voltage' 或 'Magnetic'
params.log_name = 'ProcessLog.txt';
params.valid_ext = '.csv';

%% ========================================================================
% 2. 选择源文件夹
%% ========================================================================
srcFolder = uigetdir(pwd, '请选择源数据文件夹');
if srcFolder == 0
    error('未选择文件夹，程序终止。');
end

fprintf('源文件夹：%s\n', srcFolder);

%% 打开日志文件
logPath = fullfile(srcFolder, params.log_name);
fid = fopen(logPath, 'w');
if fid == -1
    error('无法创建日志文件：%s', logPath);
end

fprintf(fid, "PolarDataManager 日志\n");
fprintf(fid, "开始时间：%s\n", datestr(now));
fprintf(fid, "DataType：%s\n", params.DataType);
fprintf(fid, "源目录：%s\n\n", srcFolder);

%% ========================================================================
% 3. 文件分类（按电场 / 磁场值）
%% ========================================================================
fprintf('>>> 正在分类文件...\n');
fprintf(fid, "步骤 1：文件分类\n");

[movedFiles, skippedFiles] = classifyFiles(srcFolder, params, fid);

fprintf(fid, "  已移动文件数：%d\n", length(movedFiles));
fprintf(fid, "  跳过文件数：%d\n\n", length(skippedFiles));

%% ========================================================================
% 4. 子目录汇总（每个数值文件夹）
%% ========================================================================
fprintf('>>> 正在生成子汇总...\n');
fprintf(fid, "步骤 2：生成子汇总\n");

aggregateByField(srcFolder, params, fid);

%% ========================================================================
% 5. 生成 Total.xlsx
%% ========================================================================
fprintf('>>> 正在生成总汇总文件...\n');
fprintf(fid, "步骤 3：总汇总\n");

createTotalSummary(srcFolder, params, fid);

fprintf(fid, "\n完成时间：%s\n", datestr(now));
fprintf(fid, "状态：成功\n");
fclose(fid);

fprintf('\n全部处理完成，总耗时 %.2f 秒。\n', toc);

% =========================================================================
% ============================ 以下为本地函数 =============================
% =========================================================================

%% ========================================================================
function [movedFiles, skippedFiles] = classifyFiles(srcFolder, params, fid)
% 识别电压/磁场值并移动文件

files = dir(fullfile(srcFolder, ['*' params.valid_ext]));
movedFiles = {}; skippedFiles = {};

if isempty(files)
    fprintf(fid, '未找到任何 CSV 文件。\n');
    return;
end

pattern = '([-+]?\d*\.?\d+)\s*([VT])';

for i = 1:length(files)
    fname = files(i).name;
    fpath = fullfile(srcFolder, fname);
    tokens = regexp(fname, pattern, 'tokens');
    fieldValue = [];

    if ~isempty(tokens)
        for t = 1:length(tokens)
            val = tokens{t}{1};
            unit = upper(tokens{t}{2});

            if strcmp(params.DataType, 'Voltage') && strcmp(unit,'V')
                fieldValue = val; break;
            elseif strcmp(params.DataType, 'Magnetic') && strcmp(unit,'T')
                fieldValue = val; break;
            end
        end
    end

    if isempty(fieldValue)
        skippedFiles{end+1} = fname;
        fprintf(fid, '  跳过（无匹配）：%s\n', fname);
        continue;
    end

    targetFolder = fullfile(srcFolder, fieldValue);
    if ~exist(targetFolder, 'dir'), mkdir(targetFolder); end

    try
        movefile(fpath, fullfile(targetFolder,fname));
        movedFiles{end+1} = fname;
        fprintf(fid, '  移动：%s  →  %s\n', fname, targetFolder);
    catch
        skippedFiles{end+1} = fname;
        fprintf(fid, '  移动失败：%s\n', fname);
    end
end
end

%% ========================================================================
function aggregateByField(srcFolder, params, fid)
% 每个子文件夹生成汇总 Excel

subdirs = getSubdirs(srcFolder);
fprintf(fid, "子目录数：%d\n", length(subdirs));

for i = 1:length(subdirs)
    sub = subdirs(i).name;
    subPath = fullfile(srcFolder, sub);

    csvs = dir(fullfile(subPath, ['*' params.valid_ext]));
    if isempty(csvs)
        fprintf(fid, "  - %s 为空，跳过\n", sub);
        continue;
    end

    dataCells = {};
    names = {};

    for j = 1:length(csvs)
        filePath = fullfile(subPath, csvs(j).name);
        data = readmatrix(filePath);

        dataCells{end+1} = data;  % 修复：保持 cell(1,N)
        [~, nameOnly, ~] = fileparts(csvs(j).name);
        names{end+1} = nameOnly;

        fprintf(fid, "  读取：%s\n", filePath);
    end

    % 修复问题：使用横向 cell2mat
    combine = cell2mat(dataCells);

    outFile = fullfile(subPath, [sub '.xlsx']);

    try
        writecell(names, outFile, 'Sheet',1, 'Range','A1');
        writematrix(combine, outFile, 'Sheet',1, 'Range','A2');
        fprintf(fid, "  汇总写入：%s\n", outFile);
    catch ME
        fprintf(fid, "  写入失败：%s\n", ME.message);
    end
end
end

%% ========================================================================
function createTotalSummary(srcFolder, ~, fid)
% 合并所有子目录汇总为 Total.xlsx

subdirs = getSubdirs(srcFolder);
fieldMap = containers.Map('KeyType','double','ValueType','any');
headerMap = containers.Map('KeyType','double','ValueType','any');

for i = 1:length(subdirs)
    sub = subdirs(i).name;
    subPath = fullfile(srcFolder, sub);

    xlsFile = fullfile(subPath, [sub '.xlsx']);
    if ~exist(xlsFile,'file')
        fprintf(fid, "  - %s 无汇总文件\n", sub);
        continue;
    end

    raw = readcell(xlsFile);
    headers = raw(1,:);
    data = cell2mat(raw(2:end,:));

    numericField = str2double(sub);
    headerMap(numericField) = headers;
    fieldMap(numericField) = data;

    fprintf(fid, "  合并：%s\n", xlsFile);
end

fields = cell2mat(fieldMap.keys);
if isempty(fields)
    fprintf(fid, "  无可汇总数据。\n");
    return;
end

fields_sorted = sort(fields);
TotalData = [];
LabelRow = {};
HeaderRow = {};

for i = 1:length(fields_sorted)
    f = fields_sorted(i);
    d = fieldMap(f);
    h = headerMap(f);

    TotalData = [TotalData, d];  % 修复：横向扩展
    LabelRow = [LabelRow, repmat({num2str(f)}, 1, size(d,2))];
    HeaderRow = [HeaderRow, h];
end

TotalTable = [
    [{' '}, LabelRow];
    [{' '}, HeaderRow];
    [cell(size(TotalData,1),1), num2cell(TotalData)];
];

out = fullfile(srcFolder, 'Total.xlsx');
writecell(TotalTable, out);

fprintf(fid, "  总汇总已写入：%s\n", out);
end

%% ========================================================================
function subdirs = getSubdirs(parentDir)
% 获取所有子目录
d = dir(parentDir);
subdirs = d([d.isdir] & ~ismember({d.name},{'.','..'}));
end
