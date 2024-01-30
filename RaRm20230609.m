% This code is used to batch calculate Ra and Rm for quality control of 
% data from whole-cell recording
clear
% Set parameters
testV=10000;
% Select seal test interval
teststart=9000;
testend=9500;
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
    Sweep=Rawdata(:,:,iSweep);
    Sealtest=Sweep(teststart:testend);
    BSL{ifile,iSweep}=Sealtest(1); % Estimate BSL
    [Peak,Peakloc]=max(Sealtest); % Calculate peak
    Plateau=Sealtest(end); % Estimate plateau
    Ra=testV/(Peak-BSL{ifile,iSweep});
    Rm=testV/(Plateau-BSL{ifile,iSweep}); 
    % Output result
    xlswrite(strcat(abfname{ifile}(1:end-4),'.xlsx'),Ra,1,strcat('A',num2str(iSweep)));
    xlswrite(strcat(abfname{ifile}(1:end-4),'.xlsx'),Rm,1,strcat('B',num2str(iSweep)));
end
end
