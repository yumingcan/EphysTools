% This code is used to batch generate the phase plot of action potentials
clear
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
v=[];
dvdt=[];
resultdata=[];
for isweep=1:1:nsweep
v=rawdata(:,:,isweep);
dvdt=gradient(v); % Calculate dvdt
% Output v and dvdt
resultdata=num2cell([v,dvdt]');
rowname={strcat('Sweep',num2str(isweep),'_v'),strcat('Sweep',num2str(isweep),'_dvdt')}';
result=[rowname,resultdata];
resultname=strcat(abfname{ifile}(1:end-4),'.xlsx');
xlswrite(resultname,result,1,num2str(2*isweep-1));
% Phase plot
dv_plot=figure();
Figname=strcat(abfname{ifile}(1:end-4),'_',num2str(isweep),'.fig');
plot(v,dvdt);
saveas(dv_plot,Figname);
end
end
close all