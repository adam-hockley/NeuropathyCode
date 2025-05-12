% sheets = sheetnames('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx');
sheets = sheetnames('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx');

data = {}; patients = {}; numpatients = [];
for i = 1:length(sheets)
    % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
    data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
    patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
    patients{i} = data{i}{patientsrows,1};
    numpatients(i) = length(patients{i});
end

figure
% tiledlayout(1,2)
% nexttile
plot(numpatients,'k','linewidth',2)

ylabel('Number of patients')
hold on
for i = 1:length(patients)
    numpatients_pre(i) = sum(ismember(patients{1},patients{i}));
end
plot(numpatients_pre,'k--','linewidth',2)
for i = 1:length(patients)
    numpatients_3(i) = sum(ismember(patients{5},patients{i}));
end
plot(numpatients_3,'k:','linewidth',2)

yline(length(unique(cell2mat(patients'))),'k:')


legend({'Patients with data','Patients with preimpant data','Patients with 3 year data','Total enrolled'},'location','best')
legend('boxoff')
beautify

xticks(1:16)
xticklabels(sheets)

% set(gcf, 'Position', [100 100 1000 500]);
print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\Counts.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\Counts.png')


%% NOW DO THE PLOTTING OVER TIME

rng("default")

sheets = sheetnames('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx');

data = {}; patients = {}; numpatients = [];
for i = 1:length(sheets)
    % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
    data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
    patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
    patients{i} = data{i}{patientsrows,1};
    numpatients(i) = length(patients{i});
end

data{8}(95:end,:) = [];
figure
tiledlayout(2,7)
nexttile
plot(numpatients,'k','linewidth',2)

ylabel('Number of patients')
hold on
for i = 1:length(patients)
    numpatients_pre(i) = sum(ismember(patients{1},patients{i}));
end
plot(numpatients_pre,'r--','linewidth',2)
for i = 1:length(patients)
    numpatients_3(i) = sum(ismember(patients{5},patients{i}));
end
plot(numpatients_3,'k:','linewidth',2)

yline(length(unique(cell2mat(patients'))),'k--')

legend({'Patients with data','Patients with preimpant data','Patients with 3 year data','Total enrolled'},'location','best')
legend('boxoff')
xticks(1:16)
xticklabels(sheets)

sheets{1} = 'PreTran';

colour = [0.2 0.2 0.2];

columns{1} = [2,6,11,15,24,32];
columns{2} = [3,7,8,12,16,17,25,33,34];
columns{3} = [4,9,13,18,26,35];
columns{4} = [5,10,14,19,27,36];
columns{5} = [20, 22, 30, 37];
columns{6} = [21, 23, 31, 38];
titles = {'LDM','VCM','DA','PA','VCS','SA'};

sigheight = ([1.4 1.4 3.5 3.5 1.4 5]);

for cl = 1:length(columns)

    % Loop
    nexttile
    output = []; outputnorm = [];
    for i = 1:length(patients{1})
        for ii = 1:length(data)
            currPatientRow = find(data{ii}{:,1} == patients{1}(i)); % Rows number for this patient
            if length(currPatientRow)<1
                output(i,ii) = nan;
                outputnorm(i,ii) = nan;
            else
                output(i,ii) = mean(data{ii}{currPatientRow,columns{cl}},'omitnan');
                currPatientRowBaseline = find(data{1}{:,1} == patients{1}(i));
                outputnorm(i,ii) = output(i,ii) / mean(data{1}{currPatientRowBaseline,columns{cl}},'omitnan'); %normalise only to these same animals at baseline
            end
        end
    end

    outputnorm(outputnorm==inf) = nan;
    [L,U,yMean] = ciprep(outputnorm(:,1:9),95);
    x = 1:9;
    % L(isnan(L)) = 0; U(isnan(U)) = 0;

    ciplot(L,U,x,colour)
    hold on
    plot(x,yMean,'r','linewidth',2)
    yline(1,'k--')
    ylabel([titles{cl} ' (norm.)'])
    % if cl == 1
    %     legend({'Group 1','','Group 2'},'location','best')
    %     legend('boxoff')
    % end
    title(titles{cl})
    xticks(1:16)
    xticklabels(sheets)

    for i = x
        [h(i) p(i)] = ttest(outputnorm(:,i)-1,0,'alpha',0.05/8);
    end
    sigparts = calcsigparts(h);
    clear h p

    plot(sigparts, ones(1,length(sigparts))*sigheight(cl),'color',colour,'linewidth', 2,'handlevisibility','off')
    
    if ismember(cl,[1 2 5])
        ylim([0.7 1.5])
    elseif ismember(cl,[3 4])
        ylim([0.5 4])
    else
        ylim([0.5 6])
    end

end



%% Now repeat for year 3 to end

nexttile

plot(numpatients,'k','linewidth',2)

ylabel('Number of patients')
hold on
for i = 1:length(patients)
    numpatients_pre(i) = sum(ismember(patients{1},patients{i}));
end
plot(numpatients_pre,'k--','linewidth',2)
for i = 1:length(patients)
    numpatients_3(i) = sum(ismember(patients{5},patients{i}));
end
plot(numpatients_3,'r:','linewidth',2)

yline(length(unique(cell2mat(patients'))),'k--')

legend({'Patients with data','Patients with preimpant data','Patients with 3 year data','Total enrolled'},'location','best')
legend('boxoff')
xticks(1:16)
xticklabels(sheets)


sheets = sheets(5:end);
data = data(5:end);
patients = patients(5:end);

% data = {}; patients = {}; numpatients = [];
% for i = 1:length(sheets)
%     % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
%     data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
%     patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
%     patients{i} = data{i}{patientsrows,1};
%     numpatients(i) = length(patients{i});
% end

colour = [0.2 0.2 0.2];

% columns{1} = [2,6,11,15,24,32];
% columns{2} = [3,7,8,12,16,17,25,33,34];
% columns{3} = [20, 22];
% columns{4} = [4,9,13,18,26,35];
% columns{5} = [5,10,14,19,27,36];
% titles = {'LDM','VCM','VCS','DA','PA'};

columns{1} = [2,6,11,15,24,32];
columns{2} = [3,7,8,12,16,17,25,33,34];
columns{3} = [4,9,13,18,26,35];
columns{4} = [5,10,14,19,27,36];
columns{5} = [20, 22, 30, 37];
columns{6} = [21, 23, 31, 38];
titles = {'LDM','VCM','DA','PA','VCS','SA'};

sigheight = ([1.4 1.4 3.5 3.5 1.4 5]);
for cl = 1:length(columns)

    % Loop
    nexttile
    output = []; outputnorm = [];
    for i = 1:length(patients{1})
        for ii = 1:length(data)
            currPatientRow = find(data{ii}{:,1} == patients{1}(i)); % Rows number for this patient
            if length(currPatientRow)<1
                output(i,ii) = nan;
                outputnorm(i,ii) = nan;
            else
                output(i,ii) = mean(data{ii}{currPatientRow,columns{cl}},'omitnan');
                currPatientRowBaseline = find(data{1}{:,1} == patients{1}(i));
                outputnorm(i,ii) = output(i,ii) / mean(data{1}{currPatientRowBaseline,columns{cl}},'omitnan'); %normalise only to these same animals at baseline
            end
        end
    end

    outputnorm(outputnorm==inf) = nan;
    [L,U,yMean] = ciprep(outputnorm(:,1:10),95);
    x = 1:10;
    % L(isnan(L)) = 0; U(isnan(U)) = 0;

    ciplot(L,U,x,colour)
    hold on
    plot(x,yMean,'r','linewidth',2)
    yline(1,'k--')
    ylabel([titles{cl} ' (norm.)'])
    % if cl == 1
    %     legend({'Group 1','','Group 2'},'location','best')
    %     legend('boxoff')
    % end
    title(titles{cl})
    xticks(1:16)
    xticklabels(sheets)

    for i = x
        [h(i) p(i)] = ttest(outputnorm(:,i)-1);
    end
    sigparts = calcsigparts(h);
    clear h p

    plot(sigparts, ones(1,length(sigparts))*sigheight(cl),'color',colour,'linewidth', 2,'handlevisibility','off')
    if ismember(cl,[1 2 5])
        ylim([0.7 1.5])
    elseif ismember(cl,[3 4])
        ylim([0.5 4])
    else
        ylim([0.5 6])
    end

end

beautify
set(gcf, 'Position', [100 100 1000 500]);
print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllChange.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllChange.png')


%%

sheets = sheetnames('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx');

data = {}; patients = {}; numpatients = [];
for i = 1:length(sheets)
    % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
    data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
    patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
    patients{i} = data{i}{patientsrows,1};
    numpatients(i) = length(patients{i});
end

titles = {'LDM Peroneal L','VCM Peroneal L','Amp Peroneal L',...
    'LDM Peroneal R','VCM Peroneal R','Amp Peroneal R',...
    'LDM Tibial L','VCM Tibial L','Amp Tibial L',...
    'LDM Tibial R','VCM Tibial R','Amp Tibial R',...
    'LDM Mediano L-M','VCM Mediano L-M','Amp Mediano L-M',...
    'LDM Cubital L-M','VCM Cubital L-M','Amp Cubital L-M',...
    'empty','VCS Sural L','Amp Sural L',...
    'empty','VCS Sural R','Amp Sural R',...
    'empty','VCM Mediano L-S','Amp Mediano L-S',...
    'empty','VCM Cubital L-S','Amp Cubital L-S'};
% columnsinclude = [20 21 22 23]; % p&t&s
columnsinclude = [6 8 9  15 17 18  2 3 4  11 12 13  24 25 26  32 33 35  1 20 21  1 22 23  1 30 31  1 37 38]; % p&t
% columnsinclude = [6 8 9 15 17 18 20 21 22 23]; % p&s
% columnsinclude = [6 8 9 15 17 18]; % p
dataBaseline = data{1}{:,columnsinclude};
dataLater = data{7}{:,columnsinclude};

% convert 0's to nan
dataBaseline(dataBaseline==0) = nan;
dataLater(dataLater==0) = nan;

figure
tiledlayout(3,10)
figureorder = [1:10:30, 2:10:30, 3:10:30, 4:10:30, 5:10:30, 6:10:30, 7:10:30, 8:10:30, 9:10:30, 10:10:30];
pall = [];
for i = 1:size(dataBaseline,2) % for each location
    % if i == 19 || i == 21
    %     nexttile
    % end
    nexttile(figureorder(i))
    unitcount = height(dataBaseline);

    % bar([mean(data1) mean(data2)],'facecolor',[0.6 0.6 0.6])
    vp(1) = violinplot([dataBaseline(:,i)], {''},'HalfViolin','left','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);
    vp(2) = violinplot([dataLater(:,i)], {''},'HalfViolin','right','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);

    pall(i) = corrected_ztest(dataBaseline(:,i),dataLater(:,i));

    % Off-center median boxes
    vp(1).MedianPlot.XData = 0.95;
    vp(2).MedianPlot.XData = 1.05;
    vp(1).BoxPlot.XData = vp(1).BoxPlot.XData-0.05;
    vp(2).BoxPlot.XData = vp(2).BoxPlot.XData+0.05;

    vp(1).ViolinColor{1} = "#bfbfbf";
    vp(2).ViolinColor{1} = "#404040";
    vp(1).ScatterPlot.MarkerFaceColor = [0 0 0];
    vp(2).ScatterPlot.MarkerFaceColor = [0.9 0.9 0.9];
    for ii = 1:2
        vp(ii).ScatterPlot.SizeData = 2;
        vp(ii).ViolinPlot.LineWidth = 1;
        vp(ii).BoxWidth = 0.05;
        vp(ii).ViolinPlot.LineWidth = 1;
        vp(ii).MeanPlot.Color = [0.5 0.5 0.5];
        vp(ii).MeanPlot.LineWidth = 2;
        vp(ii).ViolinPlot.FaceAlpha = 1;
        vp(ii).MedianPlot.SizeData = 25;  
    end
 
    if ~rem(i,3)
        ylim([0 20])
    elseif ~rem(i+1,3)
        ylim([0 80])
    else
        ylim([0 10])
    end

    if i == 3
        legend({'Pre','','','','','','','','5 yr'},'location','best')
        legend('boxoff')
    end
    labels = ({'Latency (ms)','Velocity (m/s)','Amplitude (mV)'});
    if ismember(i,[1 2 3])
        ylabel(labels{i})
    else
        set(gca,'YTickLabel',[]);
    end
    if ismember(i,[1:3:30])
        title(titles{i+1}(4:end))
    end

    xlim([0.5 1.5])
    hold on
end
FDR = mafdr(pall);
for i = 1:size(dataBaseline,2) % for each location
    nexttile(figureorder(i))
    if FDR(i) < 0.05 %% better to do false discovery or something?
        sigstar([0.9 1.1])
    end
end
beautify
set(gcf, 'Position', [100 100 1000 500]);
print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllViolins.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllViolins.png')



%%

sheets = sheetnames('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx');

sheets = sheets(5:end);
data = {}; patients = {}; numpatients = [];
for i = 1:length(sheets)
    % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
    data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
    patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
    patients{i} = data{i}{patientsrows,1};
    numpatients(i) = length(patients{i});
end

titles = {'LDM Peroneal L','VCM Peroneal L','Amp Peroneal L',...
    'LDM Peroneal R','VCM Peroneal R','Amp Peroneal R',...
    'LDM Tibial L','VCM Tibial L','Amp Tibial L',...
    'LDM Tibial R','VCM Tibial R','Amp Tibial R',...
    'LDM Mediano L-M','VCM Mediano L-M','Amp Mediano L-M',...
    'LDM Cubital L-M','VCM Cubital L-M','Amp Cubital L-M',...
    'empty','VCS Sural L','Amp Sural L',...
    'empty','VCS Sural R','Amp Sural R',...
    'empty','VCM Mediano L-S','Amp Mediano L-S',...
    'empty','VCM Cubital L-S','Amp Cubital L-S'};
% columnsinclude = [20 21 22 23]; % p&t&s
columnsinclude = [6 8 9  15 17 18  2 3 4  11 12 13  24 25 26  32 33 35  1 20 21  1 22 23  1 30 31  1 37 38]; % p&t
% columnsinclude = [6 8 9 15 17 18 20 21 22 23]; % p&s
% columnsinclude = [6 8 9 15 17 18]; % p
dataBaseline = data{1}{:,columnsinclude};
dataLater = data{7}{:,columnsinclude};

% convert 0's to nan
dataBaseline(dataBaseline==0) = nan;
dataLater(dataLater==0) = nan;

figure
tiledlayout(3,10)
figureorder = [1:10:30, 2:10:30, 3:10:30, 4:10:30, 5:10:30, 6:10:30, 7:10:30, 8:10:30, 9:10:30, 10:10:30];
for i = 1:size(dataBaseline,2) % for each location
    % if i == 19 || i == 21
    %     nexttile
    % end
    nexttile(figureorder(i))
    unitcount = height(dataBaseline);

    % bar([mean(data1) mean(data2)],'facecolor',[0.6 0.6 0.6])
    vp(1) = violinplot([dataBaseline(:,i)], {''},'HalfViolin','left','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);
    vp(2) = violinplot([dataLater(:,i)], {''},'HalfViolin','right','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);

    % Move medians
    vp(1).MedianPlot.XData = 0.95;
    vp(2).MedianPlot.XData = 1.05;
    vp(1).BoxPlot.XData = vp(1).BoxPlot.XData-0.05;
    vp(2).BoxPlot.XData = vp(2).BoxPlot.XData+0.05;

    vp(1).ViolinColor{1} = "#9B9BC5";
    vp(2).ViolinColor{1} = "#C59B9A";
    for ii = 1:2
        vp(ii).ScatterPlot.SizeData = 5;
        vp(ii).ViolinPlot.LineWidth = 1;
        vp(ii).BoxWidth = 0.05;
        vp(ii).ViolinPlot.LineWidth = 1;
        vp(ii).MeanPlot.Color = [0.5 0.5 0.5];
        vp(ii).MeanPlot.LineWidth = 2;
        vp(ii).ScatterPlot.MarkerFaceColor = [0 0 0];
        vp(ii).ViolinPlot.FaceAlpha = 0.6;
    end
 
    if ~rem(i,3)
        ylim([0 20])
    elseif ~rem(i+1,3)
        ylim([0 80])
    else
        ylim([0 10])
    end

    if i == 3
        legend({'3 yr','','','','','','','','9 yr'},'location','best')
        legend('boxoff')
    end
    labels = ({'LDM','VCM','Amp'});
    if ismember(i,[1 2 3])
        ylabel(labels{i})
    else
        set(gca,'YTickLabel',[]);
    end
    if ismember(i,[1:3:30])
        title(titles{i+1}(4:end))
    end

    xlim([0.5 1.5])
    hold on
end
beautify
set(gcf, 'Position', [100 100 1000 500]);
print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllViolins3yr.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\AllViolins3yr.png')


