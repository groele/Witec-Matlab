% -------------------------------------------------------------------------
%% === WITec 光谱导出 + 分类脚本（含 PL 二级分类 + 统计汇总） ===
% 脚本名: WITec_Spectrum_Export_and_Classification.m
%
% 功能:
% 1. 批量导出 WITec 文件（.wip, .wid）中的光谱数据，仅保存 Y 轴数据到独立的 CSV 文件。
% 2. 导出的光谱会根据其类型（如 PL, Raman）自动分类到对应的子文件夹下。
% 3. 对 PL 光谱进行更精细的二级分类，分为功率、磁场、温度和圆偏振依赖等。
% 4. 自动生成一个 Excel 统计表，汇总每种类型光谱的数量。
% 5. 移除日志功能，减少不必要的磁盘写入，从而提高脚本运行效率。
%
%%  1、PL 光谱识别规则 
% 脚本通过以下关键字识别光谱类型，请确保文件名遵循这些规则：
% 1、 功率依赖的命名规则：文件名中包含 '0.**' 格式的字符串。
%    示例: PL Sample 1 0.01
% 2、 磁场依赖的命名规则：文件名中包含数字后跟 'T' 的字符串。
%    示例: PL Sample 1 1 T
% 3、 温度依赖的命名规则：文件名中包含数字后跟 'K' 的字符串。
%    示例: PL Sample 1 30 K
% 4、 圆偏振依赖的命名规则：文件名中包含 '45' 或 '-45' 的字符串。
%    示例: PL 45 0
%
%% 2、 其他光谱识别规则 
% 保留"Raman"关键字做光谱识别
% 保留"absorb"关键字做光谱识别
% 保留'Series'关键字做光谱识别
% 保留'Spectrum'关键字做光谱识别
% 未包含上述关键字的光谱将自动保存到"Others"文件夹
%
%% 注意: 该脚本目前不处理线偏振和 Mapping 类型的数据。
%        该脚本目前未读取到 X 轴数据。
%
%% 基于关键字符进行规则匹配
%
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------

clear; clc; close all;

try
    %% 1、检查 WITio 工具箱
    if ~exist('WITio','file')
        error('WITio 工具箱未找到！请确保已正确安装，并已添加到 MATLAB 路径中。');
    end
    fprintf('=== WITec 光谱导出 + 分类工具（带详细日志 + 一级文件夹日期 + 空文件夹清理）===\n');

    %% 2、选择文件
    [filename, pathname] = uigetfile({'*.wip;*.wid','WITec 文件 (*.wip, *.wid)'}, ...
                                     '选择要导出的 WITec 文件', 'MultiSelect', 'on');
    if isequal(filename,0)
        fprintf('未选择文件，操作已取消。\n');
        return;
    end

    % 统一为 cell 数组
    if ischar(filename)
        files = {fullfile(pathname, filename)};
        file_names = {filename};
    else
        files = cellfun(@(x) fullfile(pathname,x), filename, 'UniformOutput', false);
        file_names = filename;
    end

    %% 3、创建导出根目录
    exportRoot = fullfile(pathname, 'Exported_Spectra');
    if ~exist(exportRoot,'dir'), mkdir(exportRoot); end

    % 当前日期
    dateStr = datestr(now,'yyyymmdd');

    %% 4、初始化进度条
    hWait = waitbar(0, '正在处理文件...', 'Name', '数据导出进度');

    % 初始化统计结果
    summary = [];

    %% 5、处理每个文件
    for i = 1:length(files)
        currentFile = files{i};
        [~, baseName, ~] = fileparts(file_names{i});

        waitbar((i-1)/length(files), hWait, ...
            sprintf('正在处理: %s (%d/%d)', baseName, i, length(files)));
        fprintf('\n正在处理文件: %s\n', baseName);

        % 一级输出目录（带日期后缀）
        fileOutputDir = prepareOutputDir(exportRoot, baseName, dateStr);

        % 一级分类
        destFolders = createMainCategories(fileOutputDir);

        % 导出光谱并一级分类
        stats = exportAndClassify(currentFile, destFolders);

        % PL 二级分类
        classifyPlSubfolders(destFolders.PL);

        % 汇总统计
        totalCount = stats.Raman + stats.PL + stats.Absorb + stats.Series + stats.Spectrum + stats.Others;
        summary = [summary; {baseName, stats.Raman, stats.PL, stats.Absorb, stats.Series, stats.Spectrum, stats.Others, totalCount}];
    end

    %% 6、生成统计表
    if ~isempty(summary)
        summaryTable = cell2table(summary, ...
            'VariableNames', {'FileName','Raman','PL','Absorb','Series','Spectrum','Others','Total'});

        if isscalar(file_names)
            [~, onlyBase, ~] = fileparts(file_names{1});
            summaryName = ['Summary_' onlyBase '.xlsx'];
        else
            summaryName = 'Summary_All.xlsx';
        end
        summaryPath = fullfile(exportRoot, summaryName);
        writetable(summaryTable, summaryPath);
        fprintf('\n📊 统计表已生成: %s\n', summaryPath);
    end

    %% 7、生成详细日志文件
    logPath = fullfile(exportRoot, ['Log_' dateStr '.txt']);
    fid = fopen(logPath,'w');
    if fid ~= -1
        fprintf(fid, 'WITec 光谱导出与分类详细日志\n');
        fprintf(fid, '生成日期: %s\n', datestr(now));
        fprintf(fid, '导出根目录: %s\n', exportRoot);
        fprintf(fid, '处理的数据文件: %d 个\n\n', length(files));
        
        % 记录所有处理的原始文件路径
        fprintf(fid, '原始数据文件列表:\n');
        for k = 1:length(files)
            fprintf(fid, '  [%d] %s\n', k, files{k});
        end
        fprintf(fid, '\n');

        for i = 1:size(summary,1)
            fprintf(fid, '------------------------------------------------------------\n');
            fprintf(fid, '源文件基名: %s\n', summary{i,1});
            fprintf(fid, '完整文件路径: %s\n', files{i});
            fprintf(fid, '一级文件夹: %s_%s\n', summary{i,1}, dateStr);
            fprintf(fid, '  Raman: %d\n', summary{i,2});
            fprintf(fid, '  PL: %d\n', summary{i,3});
            fprintf(fid, '  Absorb: %d\n', summary{i,4});
            fprintf(fid, '  Series: %d\n', summary{i,5});
            fprintf(fid, '  Spectrum: %d\n', summary{i,6});
            fprintf(fid, '  Others: %d\n', summary{i,7});
            fprintf(fid, '  Total spectra: %d\n', summary{i,8});

            % PL 二级分类统计
            plFolder = fullfile(exportRoot, [summary{i,1} '_' dateStr], '1、PL');
            if exist(plFolder,'dir')
                plSubs = dir(plFolder); plSubs = plSubs([plSubs.isdir]);
                fprintf(fid,'  PL 二级分类:\n');
                for j = 1:length(plSubs)
                    subname = plSubs(j).name;
                    if ~ismember(subname,{'.','..'})
                        subpath = fullfile(plFolder, subname);
                        filesInSub = dir(subpath);
                        filesInSub = filesInSub(~[filesInSub.isdir]);
                        fprintf(fid,'    %s: %d 文件\n', subname, length(filesInSub));
                    end
                end
            end
            fprintf(fid,'\n');
        end
        fclose(fid);
        fprintf('📄 详细日志已生成: %s\n', logPath);
    else
        warning('无法生成日志文件: %s', logPath);
    end

    %% 8、清理空文件夹（递归删除所有空目录）
    fprintf('\n🧹 正在检查并删除空文件夹...\n');
    removeEmptyFolders(exportRoot);
    fprintf('✅ 空文件夹清理完成。\n');

    % 关闭进度条
    if exist('hWait','var') && isgraphics(hWait)
        close(hWait);
    end

catch ME
    if exist('hWait','var') && isgraphics(hWait)
        close(hWait);
    end
    rethrow(ME);
end

%% ===== 辅助函数 =====

% 一级文件夹（带日期后缀）
function fileOutputDir = prepareOutputDir(exportRoot, baseName, dateStr)
    fileOutputDir = fullfile(exportRoot, [baseName '_' dateStr]); 
    if exist(fileOutputDir,'dir')
        rmdir(fileOutputDir,'s');
    end
    mkdir(fileOutputDir);
end

% 一级分类文件夹
function destFolders = createMainCategories(fileOutputDir)
    destFolders.PL       = fullfile(fileOutputDir,'1、PL');
    destFolders.Raman    = fullfile(fileOutputDir,'2、Raman');
    destFolders.Absorb   = fullfile(fileOutputDir,'3、Absorb');
    destFolders.Series   = fullfile(fileOutputDir,'4、Series');
    destFolders.Spectrum = fullfile(fileOutputDir,'5、Spectrum');
    destFolders.Others   = fullfile(fileOutputDir,'6、Others');

    fields = fieldnames(destFolders);
    for f = 1:numel(fields)
        mkdir(destFolders.(fields{f}));
    end
end

% 导出光谱并一级分类
function stats = exportAndClassify(currentFile, destFolders)
    stats = struct('Raman',0,'PL',0,'Absorb',0,'Series',0,'Spectrum',0,'Others',0);
    try
        [oWid, ~, ~] = WITio.read(currentFile,'-all');
    catch
        warning('无法读取文件: %s', currentFile);
        return;
    end
    if isempty(oWid), return; end

    categoryOrder = {'pl','raman','absorb','series','spectrum','others'};
    categoryFields = {'PL','Raman','Absorb','Series','Spectrum','Others'};

    for j = 1:length(oWid)
        try
            dataObj = oWid(j);
            if ~strcmp(dataObj.Type,'TDGraph'), continue; end
            % 清理文件名中的非法字符
            dataName = regexprep(dataObj.Name,'[\\/:*?"<>|]','_');
            if isempty(dataName), dataName = sprintf('spectrum_%d', j); end
            processedData = squeeze(dataObj.Data);
            if isempty(processedData), continue; end

            % 保存 Y 轴数据
            tempPath = fullfile(fileparts(destFolders.PL), [dataName '.csv']);
            if isvector(processedData)
                writematrix(processedData(:), tempPath);
            elseif size(processedData,2)>=2
                writematrix(processedData(:,2), tempPath);
            end

            fnameLower = lower(dataName);
            matched = false;
            for k = 1:numel(categoryOrder)
                if contains(fnameLower, categoryOrder{k})
                    movefile(tempPath, destFolders.(categoryFields{k}));
                    stats.(categoryFields{k}) = stats.(categoryFields{k}) + 1;
                    matched = true;
                    break;
                end
            end
            if ~matched
                movefile(tempPath, destFolders.Others);
                stats.Others = stats.Others + 1;
            end
        catch meInner
            warning('处理光谱 "%s" 时出错: %s', dataObj.Name, meInner.message);
        end
    end
end

% PL 二级分类
function classifyPlSubfolders(subFolder)
    filesSub = dir(subFolder);
    filesSub = filesSub(~[filesSub.isdir]);
    if isempty(filesSub), return; end

    folders2 = { ...
        'a、45_deg', ...
        'b、-45_deg', ...
        'c、Power dependence', ...
        'd、Voltage dependence', ...
        'e、Magnetic field', ...
        'f、Temperature', ...
        'g、Others'};

    destPaths = cellfun(@(x) fullfile(subFolder,x), folders2,'UniformOutput',false);
    cellfun(@(x) ~exist(x,'dir') && mkdir(x), destPaths);

% PL 二级分类示例脚本（优化版）
% 功能：
% 1. 根据文件名判断文件类型（45/-45度圆偏振、功率、电压、磁场、温度等）
% 2. 将文件移动到对应文件夹
% 3. 提高匹配准确性，避免误匹配

    for k = 1:length(filesSub)
        fname = filesSub(k).name;          % 当前文件名
        fpath = fullfile(subFolder, fname);% 完整路径
        matched = false;                   % 标记是否已分类
    
        %% ------------------------------
        % 1️⃣ 匹配 45度圆偏振（正向）
        % - (?<![-\d])：确保前面不是负号或数字（防止匹配 -45 或 145）
        % - 45：匹配数字 45
        % - (?!\d)：确保后面不是数字（防止匹配 450）
        %% ------------------------------
        if ~isempty(regexp(fname, '(?<![-\d])45(?!\d)', 'once'))
            safeMove(fpath, destPaths{1});
            matched = true;
    
        %% ------------------------------
        % 2️⃣ 匹配 -45度圆偏振（负向）
        % - (?<!\d) 确保前面不是数字
        % - (?!\d) 确保后面不是数字
        %% ------------------------------
        elseif ~isempty(regexp(fname, '(?<!\d)-45(?!\d)', 'once'))
            safeMove(fpath, destPaths{2});
            matched = true;
    
        %% ------------------------------
        % 3️⃣ 匹配功率依赖（0.x形式）
        % - 0\.\d+：匹配 0 开头的小数
        % - (?![kKtT])：确保后面不是 K/T 单位，避免误判温度或磁场
        %% ------------------------------
        elseif ~isempty(regexp(fname, '0\.\d+(?![kKtT])', 'once'))
            safeMove(fpath, destPaths{3});
            matched = true;
    
        %% ------------------------------
        % 4️⃣ 匹配电压依赖（如 5V, 10.5V）
        % - \b：单词边界，避免匹配 123Vabc
        % - \d+(\.\d+)?：整数或小数
        % - \s*：允许空格
        % - (?i)v：不区分大小写匹配 V
        %% ------------------------------
        elseif ~isempty(regexp(fname, '\b\d+(\.\d+)?\s*(?i)v\b', 'once'))
            safeMove(fpath, destPaths{4});
            matched = true;
    
        %% ------------------------------
        % 5️⃣ 匹配磁场依赖（如 1T, 5T）
        % - 同电压逻辑
        %% ------------------------------
        elseif ~isempty(regexp(fname, '\b\d+\s*(?i)t\b', 'once'))
            safeMove(fpath, destPaths{5});
            matched = true;
    
        %% ------------------------------
        % 6️⃣ 匹配温度依赖（如 77K, 300K）
        % - 同电压逻辑
        %% ------------------------------
        elseif ~isempty(regexp(fname, '\b\d+\s*(?i)k\b', 'once'))
            safeMove(fpath, destPaths{6});
            matched = true;
        end
    
        %% ------------------------------
        % 7️⃣ 未匹配的文件归入"其他"类别
        %% ------------------------------
        if ~matched
            safeMove(fpath, destPaths{7});
        end
    end
    
    %% ==============================
    % 使用示例：
    % 假设文件夹包含以下文件名：
    %   'Sample_45.txt'       → 分类到 destPaths{1}（45度圆偏振）
    %   'Sample_-45.txt'      → 分类到 destPaths{2}（-45度圆偏振）
    %   'Power_0.1.txt'       → 分类到 destPaths{3}（功率依赖）
    %   'Voltage_5V.txt'      → 分类到 destPaths{4}（电压依赖）
    %   'Field_1T.txt'        → 分类到 destPaths{5}（磁场依赖）
    %   'Temp_77K.txt'        → 分类到 destPaths{6}（温度依赖）
    %   'OtherFile.txt'       → 分类到 destPaths{7}（其他）
    %% ==============================

end

% 安全移动文件
function safeMove(src,dest)
    try
        movefile(src,dest);
    catch
        warning('移动文件失败: %s -> %s', src, dest);
    end
end

% 递归删除空文件夹
function removeEmptyFolders(parentDir)
    dirs = dir(parentDir);
    dirs = dirs([dirs.isdir]);
    dirs = dirs(~ismember({dirs.name}, {'.','..'}));

    for i = 1:length(dirs)
        thisDir = fullfile(parentDir, dirs(i).name);
        removeEmptyFolders(thisDir); % 递归处理子文件夹

        % 检查当前文件夹是否为空
        contents = dir(thisDir);
        contents = contents(~ismember({contents.name}, {'.','..'}));
        if isempty(contents)
            try
                rmdir(thisDir);
                % fprintf('🗑️ 已删除空文件夹: %s\n', thisDir);
            catch
                warning('删除空文件夹失败: %s', thisDir);
            end
        end
    end
end
