% -------------------------------------------------------------------------
% FileMoverByKeyword.m - 文件名关键词分类移动脚本（通用版）
% 功能：
% 遍历指定文件夹内的 Excel/CSV 文件，根据用户自定义关键词，
% 将文件分别移动到对应的子文件夹中，保证不同关键词文件互不干扰。
%
% 使用说明：
% 1. 运行程序，弹出文件夹选择对话框，选择要处理的源文件夹。
% 2. 在代码中设置 keywords 数组（可根据需要自定义多个关键词，顺序决定匹配优先级）。
% 3. 程序将自动在源文件夹下创建对应子文件夹，并将匹配的文件移动到对应文件夹。
% 4. 支持文件类型：.xls, .xlsx, .csv
% 5. 每移动一个文件会在命令行打印操作信息。
%
% 注意事项：
% - 关键词顺序决定优先匹配，避免子字符串冲突（例如 "-1.5" 和 "1.5"）。
% - 文件夹名称自动处理 "-" 替换为 "neg"，"." 替换为 "_"，以保证文件夹名合法。
%
%% ---------------------------
% 作者: Shikun Hou 
% 版本: 13.0
% 更新日期: 2025.10.09
%% ---------------------------
% -------------------------------------------------------------------------


%% 1. 选择源文件夹
sourceFolder = uigetdir(pwd,'请选择要处理的文件夹');
if isequal(sourceFolder,0)
    errordlg('未选择文件夹，程序终止','错误');
    return;
end
fprintf('已选择源文件夹: %s\n', sourceFolder);

%% 2. 用户自定义关键词
% 用户可修改以下数组，自定义需要匹配的关键词
% 注意：顺序决定优先匹配
keywords = {'-1.5', '1.5'};   
folderNames = {};              % 对应子文件夹名称

% 根据关键词生成安全文件夹名，并创建文件夹
for k = 1:length(keywords)
    safeName = strrep(keywords{k}, '-', 'neg'); % 替换负号
    safeName = strrep(safeName, '.', '_');     % 替换点
    folderNames{k} = fullfile(sourceFolder, [safeName '_files']);
    
    if ~exist(folderNames{k}, 'dir')
        mkdir(folderNames{k});
        fprintf('已创建文件夹: %s\n', folderNames{k});
    end
end

%% 3. 获取源文件夹内文件列表
fileList = dir(fullfile(sourceFolder, '*.*'));
for i = 1:length(fileList)
    [~, name, ext] = fileparts(fileList(i).name);
    
    % 跳过文件夹
    if fileList(i).isdir
        continue;
    end
    
    % 只处理 Excel / CSV 文件
    if ~any(strcmpi(ext, {'.xls', '.xlsx', '.csv'}))
        continue;
    end
    
    sourceFile = fullfile(sourceFolder, fileList(i).name);
    
    %% 4. 按关键词顺序匹配并移动文件
    for k = 1:length(keywords)
        if contains(name, keywords{k})
            destFile = fullfile(folderNames{k}, fileList(i).name);
            try
                movefile(sourceFile, destFile);
                fprintf('已移动文件: %s -> %s\n', fileList(i).name, folderNames{k});
            catch ME
                warning('无法移动文件 %s: %s', fileList(i).name, ME.message);
            end
            break; % 匹配到一个关键词后跳出
        end
    end
end

%% 5. 完成提示
disp('✅ 文件移动完成！');
