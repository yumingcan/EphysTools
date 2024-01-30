% This code is used to batch analyze spike adaptation automatically
% Data requirement: stimulus artifact amplitude must be lower than EPSC amplitude
clear
% Set parameters
StimFre=20;
Sampling=10000;
StimInterval=round(Sampling/StimFre);
nStim=StimFre;
BSLStart=1200;
BSLEnd=1400;
CalStart=1600;
PkStart=1;
PkEnd=200;

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

for iSweep=1:1:nSweep
    Rawsweep=Rawdata(:,:,iSweep);
    % Adjust raw data
    BSLdata=Rawsweep(BSLStart:BSLEnd,:,:);
    BSL=mode(BSLdata);
    Normsweep=Rawsweep-BSL;
    % Calculate EPSC
    for iStim=1:1:nStim;
    EPSCdata=Normsweep(CalStart+(iStim-1)*StimInterval+1:CalStart+iStim*StimInterval); 
    Pkdata=EPSCdata(PkStart:PkEnd);
    [sweeppks{iSweep,iStim},sweeplocs{iSweep,iStim}]=min(Pkdata);
    end    
end
    sweeppksmat=cell2mat(sweeppks);
    [nRow,nColumn]=size(sweeppksmat);
for iRow=1:1:nRow
    NormAmp(iRow,:)=100*(sweeppksmat(iRow,:)/sweeppksmat(iRow,1));
end
    NormAmpAve=mean(NormAmp,1);
    NormAmpSEM=std(NormAmp,0,1)/sqrt(nColumn);
    
% Output Result
    Numsweep=(1:nSweep);
    RowName{1}='Sweep/Stim';
for isweep=1:1:nSweep
    SweepNum=mat2str(Numsweep(isweep));
    RowName{isweep+1}=strcat('Sweep',SweepNum);
end
    RowName{nSweep+2}='Average';
    RowName{nSweep+3}='SEM';
    Num=(1:nStim);
for i=1:1:nStim
    StimNum=mat2str(Num(i));
    StimName{i}=strcat('Stim',StimNum);
end
    Summary=[NormAmp;NormAmpAve;NormAmpSEM];
    Summary_cell=num2cell(Summary);
    Resultcol=[StimName;Summary_cell];
    Result=[RowName',Resultcol];
    ResultName=strcat(abfname{ifile}(1:end-4),'.xlsx');
    xlswrite(ResultName,Result);
end

