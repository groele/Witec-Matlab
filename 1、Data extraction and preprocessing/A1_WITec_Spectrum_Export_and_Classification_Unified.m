% =========================================================================
% WITec_Spectrum_Export_and_Classification (Unified Framework v1.3)
%
% 功能：
%   1. 批量读取 WITec 文件 (.wip / .wid) 并导出所有 TDGraph 光谱 Y 轴数据为 CSV
%   2. 自动进行一级分类（PL, Raman, Absorb, Series, Spectrum, Others）
%   3. 对 PL 光谱进行二级分类（45°、-45°、功率、电压、磁场、温度、Others）
%   4. 自动生成 Summary Excel 文件
%   5. 自动生成详细日志
%   6. 自动清理空文件夹
%   7. 支持批处理模式 Batch_Processing = "True"
%
% 作者：Shikun Hou
% 版本：Unified Framework v1.3
% 更新日期：2025-11-27
% =========================================================================

clc; clear; close all;
tic;

%% ============================
% 0. 参数
% ============================
Batch_processing = "False";   % ★ 改这里：True=自动批处理整个目录；False=手动选择文件
valid_ext = {'.wip','.wid'};
date_str = datestr(now,'yyyymmdd');

PL_subfolders = { ...
    'a_45deg', ...
    'b_neg45deg', ...
    'c_power', ...
    'd_voltage', ...
    'e_magnetic', ...
    'f_temperature', ...
    'g_others'};

fprintf('=== WITec 光谱导出 + 分类工具 (Unified Framework v1.3) ===\n');

%% ============================
% 1. 选择文件 / 批处理扫描路径
% ============================
if Batch_processing == "True"
    fprintf('Batch Processing 模式开启：将自动扫描整个目录。\n');

    root_path = uigetdir(pwd,'选择要扫描的文件夹');
    if root_path == 0
        error('未选择文件夹。');
    end

    files = scanAllWITecFiles(root_path, valid_ext);
    file_names = cellfun(@(x) getFileShortName(x), files, 'UniformOutput', false);

    fprintf('共检测到 %d 个 WITec 文件。\n', length(files));

else
    [files, file_names, root_path] = selectWITecFiles(valid_ext);
    if isempty(files)
        error('未选择文件。');
    end
end

export_root = fullfile(root_path, 'Exported_Spectra');
if ~exist(export_root,'dir'), mkdir(export_root); end

summary = {};

%% ============================
% 2. 创建进度条
% ============================
hWait = waitbar(0,'初始化...');

%% ============================
% 3. 主循环：处理所有文件
% ============================
for i = 1:length(files)
    
    % 若 waitbar 句柄丢失（WITio 会刷新 UI），自动重建
    if ~exist('hWait','var') || ~isgraphics(hWait)
        hWait = waitbar(0,'处理中...');
    end

    [~, base_name, ~] = fileparts(file_names{i});

    waitbar((i-1)/length(files), hWait, ...
        sprintf('处理 %s (%d/%d)...', base_name, i, length(files)));

    fprintf('\n>>> 正在处理文件：%s\n', base_name);

    % 一级输出目录
    out_dir = createPrimaryFolder(export_root, base_name, date_str);

    % 主分类目录
    C = createMainCategoryFolders(out_dir);

    % 一级分类
    stats = exportAndPrimaryClassify(files{i}, C);

    % PL 二级分类
    classifyPLsub(C.PL, PL_subfolders);

    total = stats.PL + stats.Raman + stats.Absorb + stats.Series + stats.Spectrum + stats.Others;

    summary(end+1,:) = { ...
        base_name, stats.Raman, stats.PL, stats.Absorb, ...
        stats.Series, stats.Spectrum, stats.Others, total};
end

if isgraphics(hWait), close(hWait); end

%% ============================
% 4. 保存统计文件
% ============================
summary_path = saveSummary(export_root, summary, file_names);
fprintf('\n统计表已生成：%s\n', summary_path);

%% ============================
% 5. 保存日志
% ============================
log_path = saveLog(export_root, summary, files, date_str);
fprintf('日志已生成：%s\n', log_path);

%% ============================
% 6. 删除空文件夹
% ============================
fprintf('\n🧹 清理空文件夹...\n');
removeEmptyFolders(export_root);
fprintf('清理完成。\n');

fprintf('\n全部处理完成，总耗时 %.2f 秒。\n', toc);


% =========================================================================
% ============================ 子函数区域 ==================================
% =========================================================================

%% 文件扫描（Batch processing）
function files = scanAllWITecFiles(root_path, valid_ext)
    files = {};
    stack = {root_path};
    while ~isempty(stack)
        current = stack{1};
        stack(1) = [];

        D = dir(current);
        for i = 1:length(D)
            if D(i).isdir && ~ismember(D(i).name,{'.','..'})
                stack{end+1} = fullfile(current, D(i).name);
            else
                [~,~,ext] = fileparts(D(i).name);
                if ismember(lower(ext), valid_ext)
                    files{end+1} = fullfile(current, D(i).name);
                end
            end
        end
    end
end

%% 选择文件（非批处理）
function [files, file_names, root_path] = selectWITecFiles(valid_ext)
    [filename, pathname] = uigetfile( ...
        {'*.wip;*.wid','WITec 文件 (*.wip, *.wid)'}, ...
        '选择 WITec 文件','MultiSelect','on');

    if isequal(filename,0)
        files = {}; file_names = {}; root_path = '';
        return;
    end

    if ischar(filename)
        files = {fullfile(pathname,filename)};
        file_names = {filename};
    else
        files = cellfun(@(x) fullfile(pathname,x), filename, 'UniformOutput', false);
        file_names = filename;
    end
    root_path = pathname;
end

%% 获取短文件名
function name = getFileShortName(path)
    [~, name, ~] = fileparts(path);
end

%% 一级目录名称
function out = createPrimaryFolder(root, base, date_str)
    out = fullfile(root, [base '_' date_str]);
    if exist(out,'dir'), rmdir(out,'s'); end
    mkdir(out);
end

%% 一级分类目录
function C = createMainCategoryFolders(root)
    C.PL       = fullfile(root, '1_PL');
    C.Raman    = fullfile(root, '2_Raman');
    C.Absorb   = fullfile(root, '3_Absorb');
    C.Series   = fullfile(root, '4_Series');
    C.Spectrum = fullfile(root, '5_Spectrum');
    C.Others   = fullfile(root, '6_Others');
    S = fieldnames(C);
    for i = 1:length(S)
        mkdir(C.(S{i}));
    end
end


%% 一级分类（采用 tmp 避免 PL→PL 的移动冲突）
function stats = exportAndPrimaryClassify(witFile, C)

    stats = struct('PL',0,'Raman',0,'Absorb',0,'Series',0,'Spectrum',0,'Others',0);

    try
        [oWid, ~, ~] = WITio.read(witFile,'-all');
    catch
        warning('无法读取文件：%s', witFile);
        return;
    end

    tmpDir = fullfile(C.PL,'__tmp__');
    if ~exist(tmpDir,'dir'), mkdir(tmpDir); end

    categories = {
        'pl','PL';
        'raman','Raman';
        'absorb','Absorb';
        'series','Series';
        'spectrum','Spectrum';
    };

    % --- 1) 导出所有光谱到临时目录 ---
    for i = 1:length(oWid)

        obj = oWid(i);

        if ~strcmp(obj.Type,'TDGraph'), continue; end
        Y = squeeze(obj.Data);
        if isempty(Y), continue; end

        name = regexprep(obj.Name,'[\\/:*?"<>|]','_');
        if isempty(name), name = sprintf('spectrum_%d',i); end

        tmpPath = fullfile(tmpDir,[name '.csv']);

        if isvector(Y)
            writematrix(Y(:), tmpPath);
        else
            writematrix(Y(:,2), tmpPath);
        end
    end

    % --- 2) 分类移动 ---
    CSVs = dir(fullfile(tmpDir,'*.csv'));

    for k = 1:length(CSVs)

        fname = CSVs(k).name;
        fpath = fullfile(tmpDir,fname);
        lname = lower(fname);

        moved = false;

        for t = 1:size(categories,1)
            if contains(lname, categories{t,1})
                movefile(fpath, fullfile(C.(categories{t,2}),fname));
                stats.(categories{t,2}) = stats.(categories{t,2}) + 1;
                moved = true;
                break;
            end
        end

        if ~moved
            movefile(fpath, fullfile(C.Others,fname));
            stats.Others = stats.Others + 1;
        end
    end

    % 删除 tmp
    try, rmdir(tmpDir,'s'); end
end

%% PL 二级分类
function classifyPLsub(PL_folder, subfolders)

    files = dir(fullfile(PL_folder,'*.csv'));
    if isempty(files), return; end

    S = cellfun(@(x) fullfile(PL_folder,x), subfolders,'UniformOutput',false);
    for i = 1:length(S)
        if ~exist(S{i},'dir'), mkdir(S{i}); end
    end

    for k = 1:length(files)

        fname = files(k).name;
        src   = fullfile(PL_folder,fname);
        moved = false;

        % 45 deg
        if ~isempty(regexp(fname,'(?<![-\d])45(?!\d)','once'))
            safeMove(src, S{1}); moved=true;

        % -45 deg
        elseif ~isempty(regexp(fname,'(?<!\d)-45(?!\d)','once'))
            safeMove(src, S{2}); moved=true;

        % Power
        elseif ~isempty(regexp(fname, '0\.\d+(?![kKtT])', 'once'))
            safeMove(src, S{3}); moved=true;

        % Voltage
        elseif ~isempty(regexp(fname,'\b\d+(\.\d+)?\s*(?i)v\b','once'))
            safeMove(src, S{4}); moved=true;

        % Magnetic
        elseif ~isempty(regexp(fname,'\b\d+\s*(?i)t\b','once'))
            safeMove(src, S{5}); moved=true;

        % Temperature
        elseif ~isempty(regexp(fname,'\b\d+\s*(?i)k\b','once'))
            safeMove(src, S{6}); moved=true;
        end

        if ~moved
            safeMove(src, S{7});
        end
    end
end

%% 安全移动
function safeMove(src,dst)
    try
        movefile(src,dst);
    catch
        warning('移动失败：%s → %s', src, dst);
    end
end

%% 写 Summary 表
function out = saveSummary(root, summary, file_names)
    T = cell2table(summary, ...
        'VariableNames',{'File','Raman','PL','Absorb','Series','Spectrum','Others','Total'});

    if length(file_names)==1
        [~,base,~] = fileparts(file_names{1});
        fname = ['Summary_' base '.xlsx'];
    else
        fname = 'Summary_All.xlsx';
    end

    out = fullfile(root,fname);
    writetable(T,out);
end

%% 写日志
function out = saveLog(root, summary, files, date_str)

    out = fullfile(root, ['Log_' date_str '.txt']);
    fid = fopen(out,'w'); if fid==-1, return; end

    fprintf(fid, "WITec Log\n");
    fprintf(fid, "Date: %s\n\n", datestr(now));

    fprintf(fid, "Files processed:\n");
    for i = 1:length(files)
        fprintf(fid, "  %s\n", files{i});
    end
    fprintf(fid,"\n");

    labels = {'Raman','PL','Absorb','Series','Spectrum','Others','Total'};

    for i = 1:size(summary,1)
        fprintf(fid, "--------------------------------------------------\n");
        fprintf(fid, "File: %s\n", summary{i,1});
        for j = 1:7
            fprintf(fid, "  %s: %d\n", labels{j}, summary{i,j+1});
        end
        fprintf(fid,"\n");
    end
    fclose(fid);
end

%% 删除空文件夹
function removeEmptyFolders(parent)
    D = dir(parent);
    D = D([D.isdir] & ~ismember({D.name},{'.','..'}));
    for i = 1:length(D)
        p = fullfile(parent,D(i).name);
        removeEmptyFolders(p);
        content = dir(p);
        content = content(~ismember({content.name},{'.','..'}));
        if isempty(content)
            try, rmdir(p); end
        end
    end
end

