% =========================================================================
% FileMoverByKeyword (Unified Framework v1.0)
%
% 功能：
%   按关键词匹配文件名，将 Excel/CSV 文件移动至不同子文件夹中。
%   适用于功率依赖、温度依赖、磁场依赖等多条件数据分类。
%
% 结构：
%   1. 参数设置（关键词列表等）
%   2. 输入文件夹选择
%   3. 根据关键词自动创建目标子文件夹
%   4. 遍历源文件夹并分类移动文件
%   5. 输出日志
%
% 支持格式：.xls, .xlsx, .csv
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

% 用户自定义关键词（按顺序优先匹配）
params.keywords = {'-5', '5'};    

% 支持类型
params.extensions = {'.xls', '.xlsx', '.csv'};

%% ========================================================================
% 2. 输入文件夹选择（IO）
%% ========================================================================
sourceFolder = uigetdir(pwd, '请选择要分类的文件夹');
if isequal(sourceFolder, 0)
    error('未选择文件夹，程序终止。');
end
fprintf('已选择源文件夹：%s\n', sourceFolder);

%% ========================================================================
% 3. 创建目标子文件夹（按关键词自动生成）
%% ========================================================================
targetFolders = createTargetFolders(sourceFolder, params);

%% ========================================================================
% 4. 收集文件并分类移动
%% ========================================================================
moveCount = classifyAndMoveFiles(sourceFolder, targetFolders, params);

%% ========================================================================
% 5. 输出结果
%% ========================================================================
fprintf('\n=== 文件移动完成 ===\n');
fprintf('总共移动文件数：%d 个\n', moveCount);
fprintf('总耗时：%.2f 秒\n', toc);


% =========================================================================
% ========================== 本文件所有本地函数 ===========================
% =========================================================================


%% -------------------------------------------------------------------------
function targetFolders = createTargetFolders(sourceFolder, params)
% createTargetFolders
% 功能：
%   根据关键词自动生成合法子文件夹名，并创建文件夹
% -------------------------------------------------------------------------
keywords = params.keywords;
targetFolders = struct();

fprintf('\n=== 创建子文件夹 ===\n');

for k = 1:length(keywords)
    key = keywords{k};

    % 生成安全合法文件夹名
    safe = key;
    safe = strrep(safe, '-', 'neg');   % 例如 -5 → neg5
    safe = strrep(safe, '.', '_');     % 例如 1.5 → 1_5
    folderName = sprintf('%s_files', safe);

    fullPath = fullfile(sourceFolder, folderName);
    targetFolders(k).keyword = key;
    targetFolders(k).folder  = fullPath;

    if ~exist(fullPath, 'dir')
        mkdir(fullPath);
        fprintf('已创建文件夹：%s\n', fullPath);
    end
end
end


%% -------------------------------------------------------------------------
function moveCount = classifyAndMoveFiles(sourceFolder, targetFolders, params)
% classifyAndMoveFiles
% 功能：
%   遍历所有文件，按关键词顺序匹配并移动至对应子文件夹
% -------------------------------------------------------------------------
moveCount = 0;

% 收集文件
fileList = dir(fullfile(sourceFolder, '*.*'));

fprintf('\n=== 开始分类并移动文件 ===\n');

for i = 1:length(fileList)
    name = fileList(i).name;

    % 跳过文件夹
    if fileList(i).isdir
        continue;
    end

    % 检查扩展名
    [~, fileNameOnly, ext] = fileparts(name);
    if ~any(strcmpi(ext, params.extensions))
        continue;
    end

    fullSource = fullfile(sourceFolder, name);

    % 按关键词顺序匹配
    for k = 1:length(targetFolders)
        key = targetFolders(k).keyword;

        if contains(fileNameOnly, key)
            destPath = fullfile(targetFolders(k).folder, name);

            try
                movefile(fullSource, destPath);
                fprintf('移动：%s → %s\n', name, targetFolders(k).folder);
                moveCount = moveCount + 1;
            catch ME
                warning('无法移动文件 %s：%s', name, ME.message);
            end

            break; % 匹配一个关键词后跳出
        end
    end
end
end
