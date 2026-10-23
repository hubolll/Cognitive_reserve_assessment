function CRA_RunChecks()
%CRA_RUNCHECKS 题库完整性与导出功能的简单自动化自检。
% 在MATLAB R2024b命令窗口运行 CRA_RunChecks。
Q=CRA_QuestionBank();
assert(numel(Q)==132,'题目总数有误');
sections=unique({Q.section},'stable');
assert(numel(sections)==9,'子问卷数量有误');
assert(isequal([Q(1).module Q(end).module],[1 6]),'模块次序有误');
expected=[8 5 5 21 18 21 6 12 36];
for i=1:numel(sections)
    assert(sum(strcmp({Q.section},sections{i}))==expected(i),...
        '问卷 %s 题数有误',sections{i});
end
A=cell(1,numel(Q));
for i=1:numel(Q)
    switch Q(i).type
        case 'choice'
            A{i}=Q(i).options{1};
        case 'text'
            A{i}='测试数据';
        case 'number'
            A{i}='0';
        case 'time'
            A{i}='23:00';
        otherwise
            error('不支持的问题类型：%s',Q(i).type);
    end
end
A{strcmp({Q.id},'S01')}='23:00';
A{strcmp({Q.id},'S02')}='10';
A{strcmp({Q.id},'S03')}='07:00';
A{strcmp({Q.id},'S04')}='8';
sub=struct('Name','自检用户','ID','TEST001','Age',50,'Sex','女');
[xlsxFile,matFile,result]=CRA_Export(Q,A,sub,tempdir);
assert(isfile(xlsxFile),'没有生成Excel');
assert(isfile(matFile),'没有生成MAT');
assert(sum(result.psqiComponents)==0,'PSQI参考分值验证未通过');
assert(result.psqiComponents(4)==0,'睡眠效率0分计算未通过');
assert(isequal(size(result.summary,2),6),'汇总列数有误');
assert(numel(result.answers)==132,'导出题数有误');
key=strcmp(result.summary(:,3),'LSNS-6总分');
assert(isequal(result.summary{find(key,1),4},0),'LSNS-6计算未通过');
key=strcmp(result.summary(:,3),'PSSS总分');
assert(isequal(result.summary{find(key,1),4},12),'PSSS计算未通过');
reloaded=load(matFile,'result');
assert(numel(reloaded.result.questions)==132,'MAT数据完整性不合格');
sheetNames=cellstr(sheetnames(xlsxFile));
assert(all(ismember({'受试者信息','逐题数据','数据宽表','评分汇总','使用说明'},sheetNames)),...
    'Excel工作表不完整');
delete(xlsxFile); delete(matFile);
fprintf('全部自检通过：6模块，9问卷，132题，Excel与MAT导出正常。\n');
end
