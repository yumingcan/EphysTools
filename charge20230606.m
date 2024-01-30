% This code is used to batch calculate the charge of EPSC or IPSC based on
% normalization to local baseline
clear
% Set parameters
RMSCol=10;
ChargeFincol=8;
Unit=20000;
Sampling=10000;
UnitTime=Unit/Sampling;

% Batch import Excel files
Fileinfo=dir(fullfile('*.xlsx'));
Fileinfo_cell=struct2cell(Fileinfo);  
Filename=Fileinfo_cell(1,:);        
[~,nfile]=size(Filename);
jxlsx=0;             
for ixlsx=1:1:nfile           
    if strfind(Filename{ixlsx},'.xlsx')
        jxlsx=jxlsx+1;
        [xlsstr{jxlsx}]=xlsread(Filename{ixlsx}); 
    end
end

for ifile=1:1:nfile;

% Caculate RMS
xlsxdata=xlsstr{1,ifile};
RMSdata=xlsxdata(:,RMSCol);
RMSdata=RMSdata(~isnan(RMSdata));
RMSBSL=mode(RMSdata);
RMSAmp=RMSdata-RMSBSL;
RMSsquare=arrayfun(@(x)x*x,RMSAmp);
RMS{1,ifile}=sqrt(mean(RMSsquare));

for iCharge_Col=2:2:ChargeFincol
rawChargedata=xlsxdata(:,iCharge_Col);
Chargedata=rawChargedata(~isnan(rawChargedata));
[nChargedata,~]=size(Chargedata);
nUnit=1;
for iUnit=1:Unit:Unit*fix(nChargedata/Unit);
Unitdata=Chargedata(iUnit:iUnit+Unit-1,:);
UnitBSL=median(Unitdata);
NormUnitdata=bsxfun(@minus,Unitdata,UnitBSL.'); % Adjust each unit data
UnitI{iCharge_Col/2,nUnit,ifile}=sum(NormUnitdata); % Caculate each unit current
nUnit=nUnit+1;
end
end

TotalUnitmat=cellfun(@sum,UnitI);
EachUnitmat=TotalUnitmat(:,:,ifile);
CellCharge{ifile}=mean(EachUnitmat(:))*UnitTime; % Caculate each cell charge
end

% Output Result
TotalCharge=cell2mat(CellCharge);
Rowname={'Filename','RMS','Charge',}';
TotalCharge_cell=num2cell(TotalCharge);
Resultdata=[Filename;RMS;TotalCharge_cell];
Result=[Rowname,Resultdata];
xlswrite('Result.xlsx',Result);


