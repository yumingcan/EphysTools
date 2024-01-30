% This code can support to batch find spikes of action potentials and 
% calculate their amplitudes automatically
clear
% Set parameters
filter=0;
% 最前几个可能不存在AP，最后几个sweep可能存在非典型的AP，故需一开始确定中间可以分析的sweep
startsweep=5;
endsweep=18;

% Batch import abf files
abf=dir(fullfile('*.abf'));       
abfstr=struct2cell(abf);  
abfname=abfstr(1,:);       
[mabfname,nabfname]=size(abfname);   
jabf=0;             
for iabf=1:1:nabfname
    if strfind(abfname{iabf},'.abf')    
        jabf=jabf+1;
        [abffile{jabf}]=abfload(abfname{iabf}); 
    end
end

for ifile=1:1:nabfname
    rawdata=abffile{ifile};
    [~,~,nsweep]=size(rawdata);
for isweep=startsweep:1:endsweep
    AP=[];
    APdist=[];
    APlocs=[];
    minVISI=[];
    sweep=rawdata(:,1,isweep);
    % Find right endpoints of ISIs
    logicmat=(sweep>filter); % num to logic
    logicdiff=diff(logicmat);
    ISIright=find(logicdiff==1)+1;
    for iind=1:1:length(ISIright)-1
       [AP{iind},APdist{iind}]=max(sweep(ISIright(iind):ISIright(iind+1)));
       minVISI{iind}=min(sweep(ISIright(iind):ISIright(iind+1))); 
    end
    [AP{length(ISIright)},APdist{length(ISIright)}]=max(sweep(ISIright(end):end));
    APdist_mat=cell2mat(APdist)';
    APlocs=ISIright+APdist_mat-1;
    % Output result
    xlswrite(strcat(abfname{ifile}(1:end-4),'.xlsx'),minVISI,1,strcat('A',num2str(isweep-startsweep+1)));
    % Plot AP
    APplot=figure();
    Figname=strcat(abfname{ifile}(1:end-4),'_sweep',num2str(isweep),'.fig');
    hold on
    plot(sweep);
    plot(APlocs,sweep(APlocs),'*r');
    hold off
    saveas(APplot,Figname);
end
end
close all