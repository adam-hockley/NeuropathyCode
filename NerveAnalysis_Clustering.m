
rng("default")
data = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet','PRETRANSPLANTE');
sheets = sheetnames('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx');
% columnsinclude = [2 3 4 6 8 9 11 12 13 15 17 18 30 37]; % p&t&s v&l&a (44)
columnsinclude = [2 3 4 6 8 9 11 12 13 15 17 18]; % p&t v&l&a (49)
% columnsinclude = [2 3 4 11 12 13]; % t v&l&a (57)
% columnsinclude = [6 8 9 15 17 18 20 21 22 23]; % p&s
% columnsinclude = [6 8 9 15 17 18]; % p
data = data{:,columnsinclude};

% Select optimal number of clusters (K value) using specified range
fh = @(X,K)(kmeans(X,K));
evaSil = evalclusters(data,fh,"Silhouette","KList",2:5);
eva = evalclusters(data,fh,"CalinskiHarabasz","KList",2:5);
clear fh
K = eva.OptimalK;
clusterIndices = eva.OptimalY;

% Display cluster evaluation criterion values
figure
tiledlayout(5,3)
nexttile
yyaxis left
bar(eva.InspectedK,eva.CriterionValues);
ylabel("Calinski-Harabasz index value");
hold on
yyaxis right
plot([nan evaSil.CriterionValues],'k','linewidth',2)
ylabel("Silhouette value");
xticks(eva.InspectedK);
xlabel("n clusters");
title("Optimal n clusters = " + num2str(K))
clear K
ax = gca;
ax.YAxis(1).Color = [0 0 0.8];
ax.YAxis(2).Color = [0 0 0];

% Calculate centroids
centroids = grpstats(data,clusterIndices,"mean");

% Display results
% Display 2D scatter plot (PCA)
nexttile
[~,score] = pca(data);
clusterMeans = grpstats(score,clusterIndices,"mean");
h = gscatter(score(:,1),score(:,2),clusterIndices);
for i = 1:numel(h)
    h(i).DisplayName = strcat("Group",h(i).DisplayName);
end
clear h i
hold on
h = scatter(clusterMeans(:,1),clusterMeans(:,2),50,"kx","LineWidth",2);
hold off
h.DisplayName = "ClusterMeans";
clear h
legend({'Group 1','Group 2','Centers'},'location','northeast')
legend('boxoff')
% title("");
xlabel("PC1");
ylabel("PC2");

nexttile
scatter3(score(clusterIndices==1,1),score(clusterIndices==1,2),score(clusterIndices==1,3),200,'.')
hold on
scatter3(score(clusterIndices==2,1),score(clusterIndices==2,2),score(clusterIndices==2,3),200,'.')
% title("");
xlabel("PC1");
ylabel("PC2");
zlabel("PC3");
scatter3(clusterMeans(:,1),clusterMeans(:,2),clusterMeans(:,3),100,"kx","LineWidth",2);

%%
titles = {'Latency Peroneal L','Velocity Peroneal L','Amplitude Peroneal L',...
    'Latency Peroneal R','Velocity Peroneal R','Amplitude Peroneal R'...
    'Latency Tibial L','Velocity Tibial L','Amplitude Tibial L',...
    'Latency Tibial R','Velocity Tibial R','Amplitude Tibial R',...
    'Velocity Sural L','Velocity Sural R'};
for i = 1:size(data,2)
    if i == 13 
        nexttile
    end
    nexttile
    unitcount = sum(~isnan(clusterIndices));
    data1 = data(clusterIndices==1,i);
    data2 = data(clusterIndices==2,i);
    if length(data2) > length(data1)
        data1(end+1:length(data2)) = nan;
    else
        data2(end+1:length(data1)) = nan;
    end
    % bar([mean(data1) mean(data2)],'facecolor',[0.6 0.6 0.6])
    % vp = violinplot([data1 data2], {'1','2'},'ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);
    vp(1) = violinplot([data1], {''},'HalfViolin','left','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);
    vp(2) = violinplot([data2], {''},'HalfViolin','right','ShowData',true,'ShowMean',true,'ShowMedian',true,'MarkerSize',5);

    % Off-center median boxes
    vp(1).MedianPlot.XData = 0.95;
    vp(2).MedianPlot.XData = 1.05;
    vp(1).BoxPlot.XData = vp(1).BoxPlot.XData-0.05;
    vp(2).BoxPlot.XData = vp(2).BoxPlot.XData+0.05;

    vp(1).ViolinColor{1} = "#0072BD";
    vp(2).ViolinColor{1} = "#D95319";

    for vl = 1:2
        % vp(vl).MeanPlot.Color = [0.5 0.5 0.5];
        % vp(vl).BoxPlot.XData = vp(vl).BoxPlot.XData + [-0.03; 0.03; 0.03; -0.03];
        % vp(vl).MeanPlot.LineWidth = 2;
        % vp(vl).ScatterPlot.MarkerFaceColor = [0 0 0];
        % vp(vl).ViolinPlot.FaceAlpha = 0.6;

        vp(vl).ScatterPlot.SizeData = 5;
        vp(vl).ViolinPlot.LineWidth = 1;
        vp(vl).BoxWidth = 0.05;
        vp(vl).ViolinPlot.LineWidth = 1;
        vp(vl).MeanPlot.Color = [0.5 0.5 0.5];
        vp(vl).MeanPlot.LineWidth = 2;
        vp(vl).ScatterPlot.MarkerFaceColor = [0 0 0];
        vp(vl).ViolinPlot.FaceAlpha = 0.8;
        vp(vl).MedianPlot.SizeData = 45;  
    end
    p = signrank(data1,data2);
    if p<0.05/3
        sigstar([0.95 1.05])
    end

    if ~rem(i,3)
        ylim([0 20])
    elseif ~rem(i+1,3)
        ylim([0 60])
    else
        ylim([0 10])
    end
    if i == 1
        legend({'Group 1','','','','','','','','Group 2'},'location','best')
        legend('boxoff')
    end
    hold on
    % errorbar([mean(data1) mean(data2)],[std(data1)/sqrt(length(data1)) std(data2)/sqrt(length(data2))], 'k','linestyle','none','linewidth',2)
    title(titles{i})
end
beautify
% set(gcf, 'Position', [100 100 400 700]);

print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\ClusterSplit.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\ClusterSplit.png')

grouppatients{1} = find(clusterIndices==1);
grouppatients{2} = find(clusterIndices==2);

%% NOW DO THE PLOTTING OVER TIME FOR EACH GROUP

data = {}; patients = {}; numpatients = [];
for i = 1:length(sheets)
    % data{i} = readtable('C:\Users\admin\Desktop\BASE DE DATOS ESTUDIOS.xlsx','Sheet',sheets{i});
    data{i} = readtable('C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\BASE DE DATOS ESTUDIOS NR0.xlsx','Sheet',sheets{i});
    patientsrows = find(sum(data{i}{:,2:end},2,'omitnan') > 0);
    patients{i} = data{i}{patientsrows,1};
    numpatients(i) = length(patients{i});
end
sheets{1} = 'PreTran';

figure
tiledlayout(2,3)
colour = {[0 0.4470 0.7410],[0.8500 0.3250 0.0980]};

columns{1} = [2,6,11,15,24,32];
columns{2} = [3,7,8,12,16,17,25,33,34];
columns{3} = [4,9,13,18,26,35];
columns{4} = [5,10,14,19,27,36];
columns{5} = [20, 22, 30, 37];
columns{6} = [21, 23, 31, 38];
titles = {'LDM','VCM','DA','PA','VCS','SA'};
sigheight = ([1.4 1.4 2 2 1.4 2.8]);

for cl = 1:length(columns)

    nexttile

    for gr = 1:2
        output = []; outputnorm = [];
        for i = 1:length(grouppatients{gr})
            for ii = 1:length(data)
                currPatientRow = find(data{ii}{:,1} == grouppatients{gr}(i)); % Rows number for this patient
                if length(currPatientRow)<1
                    output(i,ii) = nan;
                    outputnorm(i,ii) = nan;
                else
                    output(i,ii) = mean(data{ii}{currPatientRow,columns{cl}},'omitnan');
                    currPatientRowBaseline = find(data{1}{:,1} == grouppatients{gr}(i));
                    outputnorm(i,ii) = output(i,ii) / mean(data{1}{currPatientRowBaseline,columns{cl}},'omitnan'); %normalise only to these same animals at baseline
                end
            end
        end

        outputnorm(outputnorm==inf) = nan;
        [L,U,yMean] = ciprep(outputnorm(:,1:9),95);
        x = 1:9;
        % L(isnan(L)) = 0; U(isnan(U)) = 0;

        ciplot(L,U,x,colour{gr})
        hold on
        plot(x,yMean,'color',colour{gr},'linewidth',2)
        yline(1,'k--')
        ylabel([titles{cl} ' (norm.)'])
        if cl == 1
            legend({'Group 1','','Group 2'},'location','best')
            legend('boxoff')
        end
        title(titles{cl})
        xticks(1:16)
        xticklabels(sheets)

        for i = x
            [h(i) p(i)] = ttest(outputnorm(:,i)-1,0,'alpha',0.05/7);
        end
        sigparts = calcsigparts(h);
        
        if gr == 1
            plot(sigparts, ones(1,length(sigparts))*sigheight(cl),'color',colour{gr},'linewidth', 2,'handlevisibility','off')
        else
            plot(sigparts, ones(1,length(sigparts))*(sigheight(cl)*1.05),'color',colour{gr},'linewidth', 2,'handlevisibility','off')
        end

        %% Save to table for mixd-effects anova
        
    end
end

sgtitle(['Per patient changes'])
beautify

print(gcf,'-vector','-depsc','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\ClusteredData.eps')
print(gcf,'-dpng','-r150','C:\Users\admin\OneDrive\Work\MATLAB\USAL Matlab\HospitalProject\ClusteredData.png')

% close all