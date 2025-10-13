% -------------------------------------------------------------------------
% 脚本功能概览：
%   FlexiExcelMerge 批量重命名、截取文件名、读取 Excel/CSV 文件并合并数据。
%   支持多种实验数据类型（功率依赖、温度依赖、磁场依赖、电压依赖）。
%
% 主要功能模块：
%   1. 文件名批量处理
%       - 自定义规则截取文件名（如 'end-4:end', '1:3'）
%       - 使用正则清除字母、空格、下划线等
%       - 自动处理重名文件
%   2. 文件读取与数据处理
%       - 支持 .csv, .xls, .xlsx 文件
%       - 自动填充缺失行（NaN）
%       - 多列文件只取第一列
%   3. 数据合并与排序
%       - 合并成一个表格，首列为 '0'
%       - 支持按列名数值升序/降序排序
%   4. 输出
%       - 保存为 MergedData.xlsx，若已存在先删除
%
% 使用注意事项：
%   - 文件名规则必须使用 MATLAB 索引格式
%   - 文件名中尽量包含数字，否则列名处理后可能为空
%   - 重复列名会在末尾加 '.1111'
%   - 多列文件仅取第一列
%
% 示例：
%   nameSliceRules = {'end-4:end'};
%   regexPattern = '[^0-9.+-]';
%   sortOption = 'reverse';
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------

clc; clear; close all;

%% -----------------------------
% 1. 用户自定义设置

nameSliceRules = {'end-4:end'};   % 默认保留末尾5个字符，可修改为多条规则
regexPattern = '[^0-9.+-]';       % 移除非数字字符
sortOption = 'forward';            % 'reverse'=降序, 'forward'=升序

%  常用场景示例：
% 
% nameSliceRulesExamples = { 
%     'end-4:end',       % 保留文件名末尾5个字符（默认）
%     '1:3',             % 保留文件名最前3个字符
%     '4:6',             % 保留文件名第4到6个字符
%     '5:end',           % 保留文件名从第5个字符到最后
%     'end-6:end-2',     % 保留倒数第7到倒数第2个字符
%     '1:2,end-1:end',   % 保留文件名前2个字符 + 最后2个字符
%     '2:5',             % 保留文件名第2到第5个字符
%     '1:2,4:5',         % 前2个字符 + 第4到5个字符
%     'end-3:end',       % 保留倒数4个字符
%     '1:3,end-2:end',   % 前3个字符 + 最后3个字符
%     'end-5:end-1',     % 保留倒数第6到倒数第2个字符
%     '3:end-2',         % 第3个字符到倒数第2个字符
% };
% 
%  使用方法示例：
%    nameSliceRules = {'end-4:end'};        % 默认末尾5个字符
%    nameSliceRules = {'1:2,end-1:end'};   % 前2个字符 + 最后2个字符

%% -----------------------------
% 2. 选择文件夹
folderPath = uigetdir(pwd,'请选择要处理的文件夹');
if isequal(folderPath,0)
    errordlg('未选择文件夹，程序终止','错误');
    return;
end
fprintf('已选择文件夹: %s\n', folderPath);

%% -----------------------------
% 3. 文件重命名
fileList = dir(fullfile(folderPath,'*.*'));
fileList = fileList(~[fileList.isdir]);

if ~isempty(fileList)
    disp('--- 批量重命名文件 ---');
    for k = 1:length(fileList)
        oldFullName = fullfile(folderPath,fileList(k).name);
        [~,basename,ext] = fileparts(fileList(k).name);

        % 截取文件名
        newName = extractNameByRule(basename, nameSliceRules);

        % 清除指定字符
        newName = regexprep(newName, regexPattern, '');

        if isempty(newName)
            warning('文件名处理后为空，跳过: %s', fileList(k).name);
            continue;
        end

        newFullName = fullfile(folderPath,[newName,ext]);

        if ~strcmp(oldFullName,newFullName) && exist(newFullName,'file')
            warning('文件 %s 已存在，跳过重命名', newFullName);
            continue;
        end

        try
            movefile(oldFullName,newFullName);
            fprintf('重命名: %s -> %s\n', fileList(k).name,[newName,ext]);
        catch ME
            warning('重命名失败: %s, 错误: %s', fileList(k).name, ME.message);
        end
    end
end
fprintf('文件重命名完成。\n');

%% -----------------------------
% 4. 删除已有汇总文件
outputFile = fullfile(folderPath,'MergedData.xlsx');
if exist(outputFile,'file')==2
    delete(outputFile);
    disp('已删除旧的 MergedData.xlsx 文件');
end

%% -----------------------------
% 5. 获取所有 Excel/CSV 文件
fileList = [dir(fullfile(folderPath,'*.csv')); ...
            dir(fullfile(folderPath,'*.xls')); ...
            dir(fullfile(folderPath,'*.xlsx'))];

if isempty(fileList)
    error('文件夹中没有 Excel/CSV 文件！');
end

%% -----------------------------
% 6. 读取数据并处理列名
mergedData = table();
maxRows = 0;
tempData = {};
validNames = {};
nameCount = containers.Map('KeyType','char','ValueType','int32');

for k = 1:length(fileList)
    fileName = fileList(k).name;
    [~,nameOnly,~] = fileparts(fileName);

    processedName = regexprep(nameOnly,'[^0-9.+-]','');
    if isempty(processedName), processedName=nameOnly; end
    if strcmp(processedName,'0'), processedName='0.0'; end
    if isKey(nameCount,processedName)
        processedName = [processedName,'.1111'];
    end
    validNames{end+1} = processedName;
    nameCount(processedName) = 1;

    try
        data = readmatrix(fullfile(folderPath,fileName));
        maxRows = max(maxRows,size(data,1));
        tempData{end+1} = data;
    catch ME
        warning('无法读取文件 %s, 错误: %s',fileName,ME.message);
    end
end

%% -----------------------------
% 7. 对齐并合并数据
for k = 1:length(tempData)
    data = tempData{k};
    nPad = maxRows - size(data,1);
    if nPad>0, data = [data; nan(nPad,size(data,2))]; end
    if size(data,2)>1, data = data(:,1); end
    currentTable = array2table(data,'VariableNames',{validNames{k}});
    mergedData = [mergedData,currentTable];
end

% 插入首列 '0'
mergedData = addvars(mergedData,nan(maxRows,1),'Before',1,'NewVariableNames',{'0'});

% 列名数值排序
allNames = mergedData.Properties.VariableNames;
numericNames = allNames(2:end);
numericValues = str2double(numericNames);
if strcmp(sortOption,'reverse'), sortDirection='descend'; else, sortDirection='ascend'; end
[~,sortIdx] = sort(numericValues,sortDirection);
sortedNames = [allNames(1), numericNames(sortIdx)];
mergedData = mergedData(:,sortedNames);

%% -----------------------------
% 8. 保存数据
try
    writetable(mergedData,outputFile);
    disp(['合并完成，文件已保存到: ',outputFile]);
catch ME
    error('保存文件失败, 错误: %s',ME.message);
end

%% -----------------------------
% 9. 文件名截取函数
function newName = extractNameByRule(basename,rules)
    newName = '';
    n = length(basename);
    for i=1:length(rules)
        rule = rules{i};
        rule = strrep(rule,'end',num2str(n));
        try
            idx = eval(['[',rule,']']);
        catch
            warning('规则 "%s" 无效，跳过',rule);
            continue;
        end
        idx(idx<1 | idx>n)=[]; 
        newName = [newName, basename(idx)];
    end
end
