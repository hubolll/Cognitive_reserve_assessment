function [xlsxPath, matPath, result] = CRA_Export(Q, A, subject, outDir)
%CRA_EXPORT  导出填写结果、附件题目评分和可复核的量表汇总。
% 在 R2024b 使用，无需 Spreadsheet Toolbox 或 Microsoft Excel。
% 不构造附件未定义的“认知储备统一总分”。
if numel(A)~=numel(Q)
    error('题库与回答数量不一致。');
end
if ~isfolder(outDir)
    error('输出文件夹不存在：%s',outDir);
end
for i=1:numel(Q)
    if strlength(strtrim(string(A{i})))==0
        error('第 %s 题尚未作答。',Q(i).id);
    end
end

% 创建不易覆盖的文件名；避免将姓名写入文件名。
idSafe=regexprep(char(string(subject.ID)),'[^A-Za-z0-9_-]','_');
if isempty(idSafe) || all(idSafe=='_'), idSafe='participant'; end
if numel(idSafe)>36, idSafe=idSafe(1:36); end
timeStamp=char(datetime('now','Format','yyyyMMdd_HHmmss'));
base0=['认知储备评估_' idSafe '_' timeStamp];
base=base0;
n=0;
while isfile(fullfile(outDir,[base '.xlsx'])) || isfile(fullfile(outDir,[base '.mat']))
    n=n+1;
    base=[base0 '_' num2str(n)];
end
xlsxPath=fullfile(outDir,[base '.xlsx']);
matPath=fullfile(outDir,[base '.mat']);

% 逐题数据：原始回答是第一手记录；来源计分仅为附件明确提供的选项权重。
cols={'模块序号','维度','子量表/文件名','题号','项目原文','原始填写/选项',...
    '选项序号','附件给定分值','活动能耗权重','加权分钟(MET-min/周)',...
    '填写单位','提示或核对事项'};
detail=cell(numel(Q)+1,numel(cols));
detail(1,:)=cols;
itemScores=nan(1,numel(Q));
for i=1:numel(Q)
    q=Q(i);
    index=NaN; score=NaN; weighted=NaN;
    if strcmp(q.type,'choice')
        index=choiceIndex(Q,A,q.id);
        if index >= 1 && index <= numel(q.scores)
            score=q.scores(index);
        end
    elseif isfinite(q.met)
        weighted=str2double(string(A{i}))*q.met;
    end
    itemScores(i)=score;
    if isnan(index), ix=''; else, ix=index; end
    if isnan(score), rawScore=''; else, rawScore=score; end
    if isnan(q.met), w=''; else, w=q.met; end
    if isnan(weighted), metscore=''; else, metscore=weighted; end
    section=q.section;
    detail(i+1,:)={q.module, getModuleName(q.module), section, q.id, q.text,...
        char(string(A{i})),ix,rawScore,w,metscore,q.unit,q.hint};
end

wideHead=[{'ID号','姓名','年龄','性别','填表时间'}, {Q.id}];
wideData=[{subject.ID,subject.Name,subject.Age,subject.Sex,...
    char(datetime('now','Format','yyyy-MM-dd HH:mm:ss'))}, A];
wide=[wideHead;wideData];

% 统计指标（量表内部规则分别解释）。
summary={'模块','子问卷','指标','结果','单位','规则或注意事项'};
summary(end+1,:)={'家庭背景','家庭背景','已填题目数',8,'题','未提供维度统一总分；父母职业ESEC需人工编码'};
summary(end+1,:)={'教育水平','教育水平','已填题目数',5,'题','未提供维度统一总分'};
summary(end+1,:)={'职业特征','职业特征','已填题目数',5,'题','职业复杂度需职业词典单独编码'};

lifestyle=0;
for i=1:13
    lifestyle=lifestyle+choiceScore(Q,A,sprintf('L%02d',i));
end
summary(end+1,:)={'生活/睡眠/运动','生活习惯','13项活动频率分值之和',lifestyle,'分（0~39）',...
    '研究性描述指标；不可解读为已验证的认知储备总分'};

[psqiTotal, psqiNames, psqiParts] = scorePSQI(Q,A);
summary(end+1,:)={'生活/睡眠/运动','睡眠习惯','PSQI参考总分',psqiTotal,'分（0~21）',...
    '按常用7成分计分；源文件部分频率标签重叠，科研使用前应核定'};
for i=1:7
    summary(end+1,:)={'生活/睡眠/运动','睡眠习惯',psqiNames{i},psqiParts(i),...
        '分（0~3）','7成分之和构成PSQI参考总分'};
end

sumAll=0; sumActive=0;
for i=1:numel(Q)
    if strcmp(Q(i).section,'运动习惯')
        amount=str2double(string(A{i}))*Q(i).met;
        sumAll=sumAll+amount;
        if ~ismember(Q(i).id,{'P06','P21'})
            sumActive=sumActive+amount;
        end
    end
end
summary(end+1,:)={'生活/睡眠/运动','运动习惯','全项目加权时间合计',sumAll,'权重×分钟/周',...
    '含“坐着”项目，非标准IPAQ活动量总分，不可直接用于临床分级'};
summary(end+1,:)={'生活/睡眠/运动','运动习惯','不含坐着的加权时间合计',sumActive,'权重×分钟/周',...
    '依据附件给出的能耗权重相乘；未校正活动重叠'};

lsns=zeros(1,6);
for i=1:6, lsns(i)=choiceScore(Q,A,sprintf('N%02d',i)); end
summary(end+1,:)={'社会交往/支持','社会交往','LSNS-6总分',sum(lsns),'分（0~30）','6道题均计入'};
summary(end+1,:)={'社会交往/支持','社会交往','LSNS-6亲属分',sum(lsns(1:3)),...
    '分（0~15）','前3项'};
summary(end+1,:)={'社会交往/支持','社会交往','LSNS-6朋友分',sum(lsns(4:6)),...
    '分（0~15）','后3项'};

psss=zeros(1,12);
for i=1:12, psss(i)=choiceScore(Q,A,sprintf('U%02d',i)); end
summary(end+1,:)={'社会交往/支持','社会支持','PSSS总分',sum(psss),'分（12~84）',...
    '12道题各1~7分'};
summary(end+1,:)={'社会交往/支持','社会支持','PSSS家庭支持',...
    sum(psss([3 4 8 11])),'分（4~28）','第3、4、8、11题'};
summary(end+1,:)={'社会交往/支持','社会支持','PSSS朋友支持',...
    sum(psss([6 7 9 12])),'分（4~28）','第6、7、9、12题'};
summary(end+1,:)={'社会交往/支持','社会支持','PSSS其他支持',...
    sum(psss([1 2 5 10])),'分（4~28）','第1、2、5、10题'};
summary(end+1,:)={'整体健康评估量表','整体健康评估量表','SF-36填写题目数',36,'题',...
    '保存36题原始选项及附件可确定的单项权重，暂不计算标准化8维得分'};
summary(end+1,:)={'整体评估','—','认知储备综合总分','未定义','—',...
    '九份附件未给出跨维度加权公式，不能擅自相加'};

person={{'字段','填写内容'};{'姓名',subject.Name};{'ID号',subject.ID};...
    {'年龄',subject.Age};{'性别',subject.Sex};...
    {'填写时间',char(datetime('now','Format','yyyy-MM-dd HH:mm:ss'))};...
    {'软件版本','1.0 / MATLAB R2024b'};{'已填题目数',numel(Q)};
    {'导出目录',outDir}};
person=vertcat(person{:});

notes={...
    '说明项目','说明内容';...
    '数据结构','“逐题数据”保留每道题的完整问题、原始答案、来源分值及活动权重；“数据宽表”每人一行便于合并分析';...
    '模块划分','按文件编号合并为6个主模块：1家庭背景、2教育水平、3职业特征、4生活/睡眠/运动、5社会交往/支持、6整体健康评估量表';...
    '认知储备总分','附件没有给出跨维度的综合计算公式；本程序不输出虚构总分或诊断';...
    'PSQI','按常见的7成分（睡眠质量、入睡时间、睡眠时长、效率、障碍、用药、日间功能）计分0～21';...
    'PSQI频率标签','附件中“1次/周”与“1-2次/周”存在重叠；显示时保留附件原文，按选项序号0～3转换';...
    'PSQI卧床时间','由晚上上床时间和次日早上起床时间按跨午夜计算；实际睡眠不得超出卧床时长';...
    '运动活动量','按附件“能量消耗权重×分钟/周”求和；含坐着的项目不能等同标准IPAQ总活动量';...
    'LSNS-6','亲属3项（0～15）+朋友3项（0～15）=总分0～30';...
    'PSSS','家庭、朋友、其他支持各4道题（4～28），总分12～84';...
    'SF-36第1题','附件未列此题换算分值：保存原始选项，来源得分留空';...
    'SF-36第8大题','身体疼痛影响工作的权重依赖第7大题，附件条件叙述需要核定，暂不自动计算';...
    'SF-36第9大题第10小题','附件给出6个选项却只列5个分值；若选第6项，来源分值留空，必须核定';...
    'SF-36维度得分','本版不计算8个标准化维度或总健康分，以免在规则不完整时给出误导分数';...
    '研究注意事项','本软件用于问卷采集与基础汇总，不提供疾病诊断或临床解释；应经研究人员核验后正式使用';...
    '隐私','文件包含姓名与ID，请放在受控的研究数据目录内，不要公开共享'};

result=struct();
result.subject=subject;
result.questions=Q;
result.answers=A;
result.itemScores=itemScores;
result.summary=summary;
result.psqiComponents=psqiParts;
result.generatedAt=datetime('now');
result.programVersion='1.0';

% xlsx文件先全部写成功才创建 MAT 数据镜像。
try
    writecell(person,xlsxPath,'Sheet','受试者信息');
    writecell(detail,xlsxPath,'Sheet','逐题数据');
    writecell(wide,xlsxPath,'Sheet','数据宽表');
    writecell(summary,xlsxPath,'Sheet','评分汇总');
    writecell(notes,xlsxPath,'Sheet','使用说明');
catch ME
    if isfile(xlsxPath)
        try, delete(xlsxPath); catch, end
    end
    rethrow(ME);
end
save(matPath,'result','-v7');
end

function name=getModuleName(n)
names={'家庭背景','教育水平','职业特征','生活/睡眠/运动',...
    '社会交往/支持','整体健康评估量表'};
name=names{n};
end

function index=choiceIndex(Q,A,id)
i=find(strcmp({Q.id},id),1);
if isempty(i), error('未知题目: %s',id); end
index=find(strcmp(Q(i).options,char(string(A{i}))),1);
if isempty(index), error('题目 %s 的选项不合法。',id); end
end

function value=choiceScore(Q,A,id)
i=find(strcmp({Q.id},id),1);
k=choiceIndex(Q,A,id);
value=Q(i).scores(k);
if isnan(value), error('题目 %s 的附件评分尚未确定。',id); end
end

function n=numberAnswer(Q,A,id)
i=find(strcmp({Q.id},id),1);
if isempty(i), error('未知题目: %s',id); end
n=str2double(string(A{i}));
if ~isfinite(n), error('题目 %s 数值不合法。',id); end
end

function t=minuteOfDay(Q,A,id)
i=find(strcmp({Q.id},id),1);
v=char(string(A{i}));
tok=regexp(v,'^(\d{1,2}):(\d{2})$','tokens','once');
if isempty(tok), error('题目 %s 时间需采用 HH:mm 格式。',id); end
h=str2double(tok{1}); mm=str2double(tok{2});
if h<0 || h>23 || mm<0 || mm>59
    error('题目 %s 时间不合法。',id);
end
t=60*h+mm;
end

function [total, labels, components] = scorePSQI(Q,A)
% 7项成分常见计分法；频次源选项1～4对应标准0～3。
subjective=choiceIndex(Q,A,'S15')-1;
latencyMin=numberAnswer(Q,A,'S02');
if latencyMin<=15, latencyClass=0;
elseif latencyMin<=30, latencyClass=1;
elseif latencyMin<=60, latencyClass=2;
else, latencyClass=3;
end
latencyScore=latencyClass + choiceIndex(Q,A,'S05')-1;
latencyComp=combinedTwo(latencyScore);

sleepHours=numberAnswer(Q,A,'S04');
if sleepHours>7, durationComp=0;
elseif sleepHours>=6, durationComp=1;
elseif sleepHours>=5, durationComp=2;
else, durationComp=3;
end
bed=minuteOfDay(Q,A,'S01');
awake=minuteOfDay(Q,A,'S03');
inBed=mod(awake-bed,1440);
if inBed==0, inBed=1440; end
if sleepHours*60 > inBed+0.01
    error('PSQI填写矛盾：实际睡眠时长超过上床至起床时间，请返回睡眠习惯检查。');
end
efficiency=100*sleepHours*60/inBed;
if efficiency>=85, efficiencyComp=0;
elseif efficiency>=75, efficiencyComp=1;
elseif efficiency>=65, efficiencyComp=2;
else, efficiencyComp=3;
end

st=0;
for i=6:14
    st=st + choiceIndex(Q,A,sprintf('S%02d',i))-1;
end
if st==0, disturbanceComp=0;
elseif st<=9, disturbanceComp=1;
elseif st<=18, disturbanceComp=2;
else, disturbanceComp=3;
end
medicine=choiceIndex(Q,A,'S16')-1;
daytime=combinedTwo(choiceIndex(Q,A,'S17')+choiceIndex(Q,A,'S18')-2);
components=[subjective latencyComp durationComp efficiencyComp disturbanceComp medicine daytime];
labels={'主观睡眠质量','入睡时间','睡眠时间','睡眠效率','睡眠障碍','催眠药物','日间功能障碍'};
total=sum(components);
end

function x=combinedTwo(v)
if v==0, x=0;
elseif v<=2, x=1;
elseif v<=4, x=2;
else, x=3;
end
end
