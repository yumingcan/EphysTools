% This code is used to batch calculate pair-pulse ratio
clear
% Set parameters
StimFre=10;
Sampling=10000;
StimInterval=1/StimFre*Sampling;
BSLStart=400;
BSLEnd=600;
CalStart=720;
PKStart=1;
PKEnd=200;

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
Rawdata=abffile{ifile};
[nPoint,nChannel,nSweep]=size(Rawdata);
sweeppks=[];
sweeplocs=[];
SweepName=[];

for iSweep=1:1:nSweep
    Rawsweep=Rawdata(:,:,iSweep);
    % Adjust raw data
    BSLdata=Rawsweep(BSLStart:BSLEnd,:,:);
    BSL=mode(BSLdata);
    Normsweep=Rawsweep-BSL;
    % Calculate EPSC
    for iStim=1:1:2;
    EPSCdata=Normsweep(CalStart+(iStim-1)*StimInterval+1:CalStart+iStim*StimInterval); 
    Pkdata=EPSCdata(PKStart:PKEnd);
    [sweeppks{iSweep,iStim},sweeplocs{iSweep,iStim}]=min(Pkdata);
    end    
end
sweeppksmat=cell2mat(sweeppks);
[nRow,nColumn]=size(sweeppksmat);
PPR=sweeppksmat(:,2)./sweeppksmat(:,1);% Calculate PPR
Resultdata=[sweeppksmat,PPR];
ResultAve=mean(Resultdata,1);
ResultSEM=std(Resultdata,0,1)/sqrt(nColumn);

% Output result
SwNum=(1:nSweep);
for iSw=1:1:nSweep
SweepNum=mat2str(SwNum(iSw));
SweepName{iSw}=strcat('Sweep',SweepNum);
end
RowName=['Sweep/Stim';SweepName';'Average';'SEM'];
StiNum=(1:2);
for iSti=1:1:2
StimNum=mat2str(StiNum(iSti));
StimName{iSti}=strcat('Stim',StimNum);
end
ColName=[StimName,'PPR'];
Summary=[Resultdata;ResultAve;ResultSEM];
Summary_cell=num2cell(Summary);
Resultcol=[ColName;Summary_cell];
Result=[RowName,Resultcol];
ResultName=strcat(abfname{ifile}(1:end-4),'.xlsx');
xlswrite(ResultName,Result);
end