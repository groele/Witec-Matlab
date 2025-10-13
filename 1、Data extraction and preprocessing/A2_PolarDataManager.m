% -------------------------------------------------------------------------
% PolarDataManager.m (V13.0) - 带日志记录版本
% 自动根据 DataType 处理电场/磁场数据
% 使用说明：
% 1. 设置数据类型 DataType 为 'Voltage' 或 'Magnetic'
%    - 'Voltage' → 处理电场数据 (*.csv 文件名包含 "V")
%    - 'Magnetic' → 处理磁场数据 (*.csv 文件名包含 "T")
% 2. 运行程序，选择源数据文件夹（包含所有 .csv 文件）。
% 3. 程序自动完成：
%    a) 按电场/磁场值生成子文件夹并移动文件
%    b) 生成每个子文件夹的汇总 Excel
%    c) 生成总汇总文件 Total.xlsx
%    d) 生成详细处理日志 ProcessLog.txt
% 4. 汇总文件内容：
%    - 第一行：原始文件名
%    - 第二行：列标题（列名）
%    - 之后行：数据
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------

clear; clc;

%% ---------------------------
%% 使用说明提示
%% ---------------------------
disp('===============================');
disp('▶️ PolarDataManager 日志版使用说明');
disp('1. 设置 DataType = ''Voltage'' 或 ''Magnetic''');
disp('2. 运行程序，选择源文件夹');
disp('3. 程序将自动完成文件分类、汇总及日志记录');
disp('===============================');

%% 0. 设置数据类型
% DataType =  Voltage   % 电场数据 ;
% DataType =  Magnetic  % 磁场数据 ;
DataType = 'Magnetic';   

switch DataType
    case 'Voltage'
        dataSuffix = 'V';  % 文件名后缀关键字
    case 'Magnetic'
        dataSuffix = 'T';
    otherwise
        error('未知数据类型，请设置 DataType 为 Voltage 或 Magnetic');
end

%% 1. 选择源文件夹
srcFolder = uigetdir(pwd, '请选择源文件夹');
if srcFolder == 0
    disp('❌ 未选择文件夹，程序终止');
    return; 
end

%% 初始化日志
logFile = fullfile(srcFolder, 'ProcessLog.txt');
fid = fopen(logFile, 'w');
if fid == -1
    error('无法创建日志文件: %s', logFile);
end

fprintf(fid, 'PolarDataManager 处理日志\n');
fprintf(fid, '开始时间: %s\n', datestr(now));
fprintf(fid, '数据类型: %s\n', DataType);
fprintf(fid, '源文件夹: %s\n\n', srcFolder);

%% 2. 执行数据处理流程
try
    fprintf('▶️ 正在进行文件分类 (%s)...\n', DataType);
    fprintf(fid, '步骤 1: 文件分类开始\n');
    [movedFiles, skippedFiles] = classifyFiles(srcFolder, DataType, fid); 
    fprintf(fid, '步骤 1: 文件分类完成\n');
    fprintf(fid, '  - 成功移动文件数: %d\n', length(movedFiles));
    fprintf(fid, '  - 跳过文件数: %d\n', length(skippedFiles));
    for i = 1:length(skippedFiles)
        fprintf(fid, '    跳过文件: %s\n', skippedFiles{i});
    end
    disp('✅ 文件分类完成！');
    
    fprintf('▶️ 正在进行磁场/电场数据汇总...\n');
    fprintf(fid, '\n步骤 2: 数据汇总开始\n');
    aggregateByField(srcFolder, fid); 
    fprintf(fid, '步骤 2: 数据汇总完成\n');
    disp('✅ 汇总完成！');
    
    fprintf('▶️ 正在生成总汇总文件...\n');
    fprintf(fid, '\n步骤 3: 生成总汇总文件开始\n');
    createTotalSummary(srcFolder, fid); 
    fprintf(fid, '步骤 3: 生成总汇总文件完成\n');
    disp('✅ 总汇总完成！');
    
    fprintf(fid, '\n处理完成时间: %s\n', datestr(now));
    fprintf(fid, '处理状态: 成功\n');
    
catch ME
    fprintf(2, '❌ 发生错误: %s\n', ME.message);
    fprintf(2, '错误发生于文件: %s, 行号: %d\n', ME.stack(1).file, ME.stack(1).line);
    fprintf(fid, '\n处理错误: %s\n', ME.message);
    fprintf(fid, '处理状态: 失败\n');
end

fclose(fid);

%% -----------------------------
%% 函数：classifyFiles
%% -----------------------------
function [movedFiles, skippedFiles] = classifyFiles(srcFolder, DataType, fid)
    files = dir(fullfile(srcFolder, '*.csv'));
    if isempty(files)
        warning('未找到任何 .csv 文件！');
        fprintf(fid, '  - 未找到任何 .csv 文件！\n');
        movedFiles = {};
        skippedFiles = {};
        return;
    end
    
    fprintf(fid, '  - 找到 %d 个 .csv 文件\n', length(files));
    
    movedFiles = {};
    skippedFiles = {};
    
    % 设置目标后缀
    switch DataType
        case 'Voltage'
            targetSuffix = 'V';
        case 'Magnetic'
            targetSuffix = 'T';
        otherwise
            error('未知数据类型，请设置 DataType 为 Voltage 或 Magnetic');
    end
    
    for i = 1:length(files)
        fname = files(i).name;
        
        % -----------------------------
        % 通用正则表达式
        % -----------------------------
        % 匹配数字 + 后缀 (V/T)，允许空格和前缀文字
        pattern = '([-+]?\d*\.?\d+)\s*([VT])';
        tokens = regexp(fname, pattern, 'tokens');
        
        fieldFolder = [];
        if ~isempty(tokens)
            % 遍历所有匹配，找到符合当前 DataType 的值
            for k = 1:length(tokens)
                val = tokens{k}{1};
                suffix = tokens{k}{2};
                if strcmp(DataType, 'Voltage') && strcmp(suffix,'V')
                    fieldFolder = val; break;
                elseif strcmp(DataType, 'Magnetic') && strcmp(suffix,'T')
                    fieldFolder = val; break;
                end
            end
        end
        
        if isempty(fieldFolder)
            msg = sprintf('未识别 %s 值，跳过文件: %s', targetSuffix, fname);
            warning(msg);
            fprintf(fid, '    - %s\n', msg);
            skippedFiles{end+1} = fname;
            continue;
        end
        
        % 创建子文件夹
        targetFolder = fullfile(srcFolder, fieldFolder);
        if ~exist(targetFolder, 'dir')
            mkdir(targetFolder);
            fprintf(fid, '    - 创建文件夹: %s\n', fieldFolder);
        end
        
        % 移动文件
        try
            sourceFilePath = fullfile(srcFolder, fname);
            targetFilePath = fullfile(targetFolder, fname);
            movefile(sourceFilePath, targetFilePath);
            fprintf(fid, '    - 移动文件: %s -> %s\n', sourceFilePath, targetFilePath);
            movedFiles{end+1} = fname;
        catch
            msg = sprintf('无法移动文件: %s', fullfile(srcFolder, fname));
            warning(msg);
            fprintf(fid, '    - %s\n', msg);
            skippedFiles{end+1} = fname;
        end
    end
end


%% -----------------------------
%% 函数：aggregateByField
%% -----------------------------
function aggregateByField(baseFolder, fid)
    fieldFolders = getSubdirs(baseFolder);
    fprintf(fid, '  - 找到 %d 个磁场/电场子文件夹\n', length(fieldFolders));
    
    for f = 1:length(fieldFolders)
        fieldPath = fullfile(baseFolder, fieldFolders(f).name);
        csvFiles = dir(fullfile(fieldPath, '*.csv'));
        if isempty(csvFiles)
            fprintf(fid, '    - 子文件夹 %s 中没有CSV文件，跳过\n', fieldFolders(f).name);
            continue;
        end
        
        fprintf(fid, '    - 处理子文件夹: %s (包含 %d 个CSV文件)\n', fieldFolders(f).name, length(csvFiles));
        
        allData = cell(1,length(csvFiles));
        fileNames = cell(1,length(csvFiles));
        
        for i = 1:length(csvFiles)
            filePath = fullfile(fieldPath, csvFiles(i).name);
            data = readmatrix(filePath);
            allData{i} = data;
            [~, nameOnly, ~] = fileparts(csvFiles(i).name);
            fileNames{i} = nameOnly;
            fprintf(fid, '      - 读取文件: %s (大小: %dx%d)\n', filePath, size(data,1), size(data,2));
        end
        
        combinedData = cell2mat(allData);
        xlsFileName = fullfile(fieldPath, [fieldFolders(f).name '.xlsx']);
        
        try
            writecell(fileNames, xlsFileName, 'Sheet', 1, 'Range', 'A1');
            writematrix(combinedData, xlsFileName, 'Sheet', 1, 'Range', 'A2');
            fprintf(fid, '      - 生成汇总文件: %s (总数据大小: %dx%d)\n', xlsFileName, size(combinedData,1), size(combinedData,2));
        catch ME
            msg = sprintf('无法写入 Excel 文件 %s: %s', xlsFileName, ME.message);
            warning(msg);
            fprintf(fid, '      - %s\n', msg);
        end
    end
end

%% -----------------------------
%% 函数：createTotalSummary
%% -----------------------------
function createTotalSummary(baseFolder, fid)
    fieldFolders = getSubdirs(baseFolder);
    fprintf(fid, '  - 从 %d 个子文件夹生成总汇总\n', length(fieldFolders));
    
    fieldDataMap = containers.Map('KeyType','double','ValueType','any');
    fieldHeaderMap = containers.Map('KeyType','double','ValueType','any');
    
    for f = 1:length(fieldFolders)
        fieldPath = fullfile(baseFolder, fieldFolders(f).name);
        xlsFile = dir(fullfile(fieldPath, '*.xlsx'));
        if isempty(xlsFile)
            fprintf(fid, '    - 子文件夹 %s 中没有Excel文件，跳过\n', fieldFolders(f).name);
            continue;
        end
        
        [~, nameOnly, ~] = fileparts(xlsFile(1).name);
        fieldNum = str2double(nameOnly);  % 只保留数字
        
        try
            xlsFilePath = fullfile(fieldPath, xlsFile(1).name);
            rawData = readcell(xlsFilePath);
            headerRow = rawData(1,:);
            data = cell2mat(rawData(2:end,:));
            fieldDataMap(fieldNum) = data;
            fieldHeaderMap(fieldNum) = headerRow;
            fprintf(fid, '    - 读取汇总文件: %s (数据大小: %dx%d)\n', xlsFilePath, size(data,1), size(data,2));
        catch ME
            xlsFilePath = fullfile(fieldPath, xlsFile(1).name);
            msg = sprintf('无法读取文件 %s: %s', xlsFilePath, ME.message);
            warning(msg);
            fprintf(fid, '    - %s\n', msg);
        end
    end
    
    fields = cell2mat(fieldDataMap.keys);
    [sortedFields, ~] = sort(fields);
    
    fprintf(fid, '  - 按磁场/电场值排序: [');
    for i = 1:length(sortedFields)
        fprintf(fid, '%g', sortedFields(i));
        if i < length(sortedFields)
            fprintf(fid, ', ');
        end
    end
    fprintf(fid, ']\n');
    
    combinedData = [];
    combinedLabels = {};
    combinedHeaders = {};
    
    for k = 1:length(sortedFields)
        fieldNum = sortedFields(k);
        data = fieldDataMap(fieldNum);
        headerRow = fieldHeaderMap(fieldNum);
        combinedData = [combinedData, data];
        combinedLabels = [combinedLabels, repmat({num2str(fieldNum)}, 1, size(data,2))];
        combinedHeaders = [combinedHeaders, headerRow];
        fprintf(fid, '    - 合并 %g 的数据 (列数: %d)\n', fieldNum, size(data,2));
    end
    
    if ~isempty(combinedData)
        combinedTable = [ ...
            [{' '}, combinedLabels]; 
            [{' '}, combinedHeaders]; 
            [cell(size(combinedData,1),1), num2cell(combinedData)];
        ];
        totalFileName = fullfile(baseFolder, 'Total.xlsx');
        try
            writecell(combinedTable, totalFileName);
            fprintf(fid, '    - 生成总汇总文件: %s (总数据大小: %dx%d)\n', totalFileName, size(combinedTable,1), size(combinedTable,2));
        catch ME
            msg = sprintf('无法写入总汇总文件 %s: %s', totalFileName, ME.message);
            warning(msg);
            fprintf(fid, '    - %s\n', msg);
        end
    else
        msg = '没有可供汇总的数据';
        fprintf(fid, '    - %s\n', msg);
        warning(msg);
    end
end

%% -----------------------------
%% 辅助函数：获取子文件夹
%% -----------------------------
function subdirs = getSubdirs(parentDir)
    allDirs = dir(parentDir);
    subdirs = allDirs([allDirs.isdir] & ~ismember({allDirs.name},{'.','..'}));
end



