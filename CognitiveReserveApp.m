function CognitiveReserveApp()
%COGNITIVERESERVEAPP 认知储备相关问卷评估/采集软件（MATLAB R2024b）。
%
% 用法：将本文件、CRA_QuestionBank.m、CRA_Export.m 放在同一文件夹，
% 在MATLAB命令窗口执行 CognitiveReserveApp。
%
% 本软件用于问卷采集和量表内的有依据计算，不计算未定义的认知储备总分。

Q=CRA_QuestionBank();
A=repmat({''},1,numel(Q));
pages=makePages(Q,5);    % 每页至多5道题，逐问卷依次展示
currentPage=0;          % 0=欢迎页；1~numel(pages)=填写页；-1=完成页
started=false;
subject=struct('Name','','ID','','Age',NaN,'Sex','');
outputDir=pwd;
visibleIndices=[];
visibleControls={};

fig=uifigure('Name','认知储备功能评估系统  |  MATLAB R2024b',...
    'Position',[110 70 1180 790], 'Color',[0.951 0.964 0.981],...
    'Resize','off', 'CloseRequestFcn',@confirmClose);
uilabel(fig,'Position',[38 718 1100 42],'Text','认知储备功能评估系统',...
    'FontSize',26,'FontWeight','bold','FontColor',[0.09 0.20 0.36]);
logoPath = fullfile(fileparts(mfilename('fullpath')), 'Slogo.png');

if isfile(logoPath)
    uiimage(fig, ...
        'ImageSource', logoPath, ...
        'Position', [745 675 550 100], ...
        'ScaleMethod', 'fit');
end
subtitle=uilabel(fig,'Position',[40 682 1100 30],...
    'Text','六大模块 · 九份问卷 · 统一采集与结果导出',...
    'FontSize',13,'FontColor',[0.40 0.48 0.58]);
content=uipanel(fig,'Position',[35 137 1110 528],...
    'BackgroundColor',[1 1 1],'BorderType','none');
backButton=uibutton(fig,'push','Text','← 上一页',...
    'Position',[40 46 148 47],'FontSize',14,...
    'ButtonPushedFcn',@(~,~)goBack());
nextButton=uibutton(fig,'push','Text','下一页 →',...
    'Position',[972 46 168 47],'FontSize',14,'FontWeight','bold',...
    'BackgroundColor',[0.15 0.38 0.69],'FontColor',[1 1 1],...
    'ButtonPushedFcn',@(~,~)goNext());
progress=uilabel(fig,'Position',[218 49 735 40],...
    'HorizontalAlignment','center','FontSize',12,...
    'FontColor',[0.36 0.45 0.55]);

renderIntro();

    function renderIntro()
        currentPage=0;
        delete(content.Children);
        backButton.Visible='off';
        nextButton.Visible='off';
        subtitle.Text='欢迎使用 · 六大模块、九份子问卷、132个项目';
        progress.Text='请先填写受试者基本信息，并选择文件输出路径';

        uilabel(content,'Position',[42 459 1020 43],...
            'Text','欢迎使用认知储备功能评估系统',...
            'FontSize',23,'FontWeight','bold','FontColor',[0.10 0.29 0.48]);
        uilabel(content,'Position',[44 390 1020 67],...
            'Text',['本系统按家庭背景、教育水平、职业特征、生活/睡眠/运动、社会交往/支持、' ...
            '整体健康评估依次采集问卷。所有项目填写完成后，将自动导出Excel逐题结果、量表汇总以及MAT数据备份。' ...
            '结果仅供研究参考，不自动作出疾病诊断或未经定义的认知储备总分。'],...
            'FontSize',13,'WordWrap','on','FontColor',[0.32 0.38 0.44]);

        addFieldLabel('受试者姓名',44,342);
        nameEdit=uieditfield(content,'text','Position',[44 300 448 34],...
            'Value',subject.Name,'FontSize',14);
        addFieldLabel('ID号（保留前导零）',571,342);
        idEdit=uieditfield(content,'text','Position',[571 300 448 34],...
            'Value',subject.ID,'FontSize',14);
        addFieldLabel('年龄（周岁）',44,263);
        ageText='';
        if isfinite(subject.Age), ageText=num2str(subject.Age); end
        ageEdit=uieditfield(content,'text','Position',[44 220 448 34],...
            'Value',ageText,'FontSize',14);
        addFieldLabel('性别',571,263);
        sexItems={'请选择','男','女','其他/不愿透露'};
        sexValue=subject.Sex;
        if isempty(sexValue), sexValue='请选择'; end
        sexSelect=uidropdown(content,'Items',sexItems,'Value',sexValue,...
            'Position',[571 220 448 34],'FontSize',14);
        addFieldLabel('结果文件输出路径',44,183);
        dirEdit=uieditfield(content,'text','Value',outputDir,...
            'Position',[44 137 815 34],'FontSize',12);
        uibutton(content,'push','Text','浏览文件夹…',...
            'Position',[872 137 148 34],'FontSize',12,...
            'ButtonPushedFcn',@(~,~)selectFolder());
        uibutton(content,'push','Text','开始评估 →',...
            'Position',[393 39 328 59], 'FontSize',18,...
            'FontWeight','bold','BackgroundColor',[0.15 0.38 0.69],...
            'FontColor',[1 1 1], 'ButtonPushedFcn',@(~,~)beginSurvey());

        function addFieldLabel(txt,x,y)
            uilabel(content,'Position',[x y 440 28],'Text',txt,...
                'FontSize',13,'FontWeight','bold');
        end

        function selectFolder()
            startPath=dirEdit.Value;
            if ~isfolder(startPath), startPath=pwd; end
            chosen=uigetdir(startPath,'选择评估结果的保存文件夹');
            if ~isequal(chosen,0), dirEdit.Value=chosen; end
        end

        function beginSurvey()
            name=strtrim(nameEdit.Value);
            pid=strtrim(idEdit.Value);
            age=str2double(strtrim(ageEdit.Value));
            sex=sexSelect.Value;
            out=strtrim(dirEdit.Value);
            if isempty(name)
                uialert(fig,'请填写受试者姓名。','基本信息未完成'); return;
            end
            if isempty(pid)
                uialert(fig,'请填写受试者ID号。','基本信息未完成'); return;
            end
            if ~isfinite(age) || age<0 || age>120 || age~=floor(age)
                uialert(fig,'年龄必须是0～120之间的整数。','基本信息有误'); return;
            end
            if strcmp(sex,'请选择')
                uialert(fig,'请选择性别。','基本信息未完成'); return;
            end
            if ~isfolder(out)
                uialert(fig,'输出路径不存在，请选择已存在的文件夹。','输出路径有误'); return;
            end
            subject=struct('Name',name,'ID',pid,'Age',age,'Sex',sex);
            outputDir=out;
            started=true;
            currentPage=1;
            renderPage();
        end
    end

    function renderPage()
        if currentPage<1 || currentPage>numel(pages), return; end
        delete(content.Children);
        pg=pages(currentPage);
        visibleIndices=pg.indices;
        visibleControls=cell(1,numel(visibleIndices));
        firstQ=Q(visibleIndices(1));
        subtitle.Text=sprintf('模块 %d / 6  ·  %s  |  当前问卷：%s（第%d/%d页）',...
            firstQ.module,moduleLabel(firstQ.module), pg.sectionName,...
            pg.subPage,pg.subTotal);
        backButton.Visible='on';
        nextButton.Visible='on';
        if currentPage==numel(pages)
            nextButton.Text='完成并导出 ✓';
        else
            nextButton.Text='下一页 →';
        end
        for k=1:numel(visibleIndices)
            index=visibleIndices(k);
            q=Q(index);
            y=526-101*k;
            panel=uipanel(content,'Position',[15 y 1080 95],...
                'BorderType','none','BackgroundColor',[0.963 0.975 0.992]);
            uilabel(panel,'Position',[20 43 1038 43],...
                'Text',[q.id '  ' q.text], 'FontSize',13,...
                'FontWeight','bold','FontColor',[0.12 0.23 0.35],...
                'WordWrap','on');
            prior=A{index};
            if strcmp(q.type,'choice')
                items=[{'— 请选择答案 —'}, q.options];
                initial='— 请选择答案 —';
                if any(strcmp(q.options,char(string(prior))))
                    initial=char(string(prior));
                end
                ctrl=uidropdown(panel,'Items',items,'Value',initial,...
                    'Position',[20 10 526 34],'FontSize',13);
            else
                ctrl=uieditfield(panel,'text','Value',char(string(prior)),...
                    'Position',[20 10 526 34],'FontSize',13);
            end
            visibleControls{k}=ctrl;
            note=strtrim([q.unit '  ' q.hint]);
            if isempty(note)
                if strcmp(q.type,'choice'), note='请选择最符合实际情况的一项';
                else, note='本项为必填'; end
            end
            uilabel(panel,'Position',[568 8 475 41],'Text',note,...
                'FontSize',11,'WordWrap','on',...
                'FontColor',[0.40 0.46 0.55]);
        end
        nAnswered=sum(cellfun(@(v)~isempty(strtrim(char(string(v)))),A));
        progress.Text=sprintf('全问卷第 %d / %d 页  |  已填写 %d / %d 项（可返回修改）',...
            currentPage,numel(pages),nAnswered,numel(Q));
    end

    function goBack()
        if currentPage<=0, return; end
        storePage(false);
        if currentPage==1
            renderIntro();
        else
            currentPage=currentPage-1;
            renderPage();
        end
    end

    function goNext()
        if currentPage<1, return; end
        [ok,warningText]=storePage(true);
        if ~ok
            uialert(fig,warningText,'请补充或检查本页答案');
            return;
        end
        if currentPage<numel(pages)
            currentPage=currentPage+1;
            renderPage();
        else
            finishSurvey();
        end
    end

    function [ok,warningText]=storePage(validate)
        ok=true;
        warningText='';
        for k=1:numel(visibleIndices)
            ix=visibleIndices(k);
            q=Q(ix);
            value=strtrim(char(string(visibleControls{k}.Value)));
            if strcmp(q.type,'choice') && strcmp(value,'— 请选择答案 —')
                value='';
            end
            A{ix}=value;
            if ~validate, continue; end
            if isempty(value)
                ok=false;
                warningText=sprintf('%s：%s\n此题必须填写。',q.id,q.text);
                return;
            end
            if strcmp(q.type,'number')
                num=str2double(value);
                if ~isfinite(num) || num<q.limits(1) || num>q.limits(2)
                    ok=false;
                    warningText=sprintf('%s：请填写有效的数值（范围 %g～%g）。',...
                        q.id,q.limits(1),q.limits(2));
                    return;
                end
            elseif strcmp(q.type,'time')
                tok=regexp(value,'^(\d{1,2}):(\d{2})$','tokens','once');
                if isempty(tok)
                    ok=false;
                    warningText=sprintf('%s：时间必须用24小时制 HH:mm，例如 22:30。',q.id);
                    return;
                end
                h=str2double(tok{1}); minute=str2double(tok{2});
                if h>23 || minute>59
                    ok=false;
                    warningText=sprintf('%s：请输入有效的24小时制时间。',q.id);
                    return;
                end
                A{ix}=sprintf('%02d:%02d',h,minute);
            elseif strcmp(q.type,'choice')
                if ~any(strcmp(value,q.options))
                    ok=false;
                    warningText=sprintf('%s：请选择有效选项。',q.id);
                    return;
                end
            end
        end
    end

    function finishSurvey()
        incomplete=find(cellfun(@(v)isempty(strtrim(char(string(v)))),A),1);
        if ~isempty(incomplete)
            indexPage=find(cellfun(@(pg)any(pg.indices==incomplete),num2cell(pages)),1);
            if ~isempty(indexPage), currentPage=indexPage; renderPage(); end
            uialert(fig,['尚有未完成项目：' Q(incomplete).id],'尚未完成');
            return;
        end
        try
            [xlsxPath,matPath]=CRA_Export(Q,A,subject,outputDir);
        catch ME
            if contains(ME.message,'PSQI填写矛盾')
                currentPage=find(strcmp({pages.sectionName},'睡眠习惯'),1);
                renderPage();
            end
            uialert(fig,ME.message,'导出失败，请核对填写内容或文件权限',...
                'Icon','error');
            return;
        end
        currentPage=-1;
        delete(content.Children);
        backButton.Visible='off';
        nextButton.Visible='off';
        subtitle.Text='评估填写完成 · 结果已成功导出';
        progress.Text='请妥善保管包含姓名和ID号的研究数据';
        uilabel(content,'Position',[85 391 945 70],...
            'Text','✓ 所有项目已完成，结果导出成功',...
            'FontWeight','bold','FontSize',23,...
            'FontColor',[0.10 0.41 0.32]);
        uilabel(content,'Position',[85 278 940 102],...
            'Text',sprintf('Excel结果文件：\n%s\n\nMAT原始数据备份：\n%s',xlsxPath,matPath),...
            'FontSize',12,'WordWrap','on');
        uilabel(content,'Position',[85 202 930 50],...
            'Text','Excel包含“受试者信息、逐题数据、数据宽表、评分汇总、使用说明”5张工作表。',...
            'FontSize',13,'WordWrap','on');
        uibutton(content,'push','Text','填写下一位受试者',...
            'Position',[375 70 356 60],'FontSize',17,...
            'BackgroundColor',[0.15 0.38 0.69],'FontColor',[1 1 1],...
            'ButtonPushedFcn',@(~,~)newSubject());
    end

    function newSubject()
        A=repmat({''},1,numel(Q));
        subject=struct('Name','','ID','','Age',NaN,'Sex','');
        started=false;
        renderIntro();
    end

    function confirmClose(~,~)
        if started && currentPage>0
            action=uiconfirm(fig,'关闭窗口会丢失尚未导出的填写内容，是否退出？',...
                '确认退出','Options',{'退出','继续填写'},...
                'DefaultOption',2,'CancelOption',2);
            if ~strcmp(action,'退出'), return; end
        end
        delete(fig);
    end
end

function pages=makePages(Q,pageSize)
% 按九份问卷的原始顺序形成页，避免把不同子量表混在一页。
sections=unique({Q.section},'stable');
pages=struct('sectionName',{},'indices',{},'subPage',{},'subTotal',{});
for s=1:numel(sections)
    inds=find(strcmp({Q.section},sections{s}));
    pageCount=ceil(numel(inds)/pageSize);
    for k=1:pageCount
        left=(k-1)*pageSize+1;
        right=min(k*pageSize,numel(inds));
        pages(end+1)=struct('sectionName',sections{s},...
            'indices',inds(left:right),'subPage',k,'subTotal',pageCount);
    end
end
end

function txt=moduleLabel(n)
switch n
    case 1, txt='家庭背景';
    case 2, txt='教育水平';
    case 3, txt='职业特征';
    case 4, txt='生活习惯 / 睡眠习惯 / 运动习惯';
    case 5, txt='社会交往 / 社会支持';
    case 6, txt='整体健康评估量表';
    otherwise, txt='';
end
end
