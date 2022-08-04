[filename,path] = uigetfile('.xlsx'); %On sélectionne l'excel à analyser

cd(path) %On se place dans le dossier contenant l'excel sélectionné
Data = readtable([path,filename]); %On ouvre les données

% On génére un excel contenant les en-têtes Nb, Amplitude et Duration, dans
% la feuille "Moyennes"
Moyennes_Name = ["Nb", "Amplitude", "Duration"]; 
writematrix(Moyennes_Name, 'Analyse_2P.xlsx', 'Sheet', 'Moyennes', 'Range', 'A1')
writematrix(Moyennes_Name, 'Analyse_2P.xlsx', 'Sheet', 'Moyennes', 'Range', 'E1')

% On demande la valeur de la balise temporelle"
Balise = inputdlg("Balise temporelle (ms) ?");
Balise = str2num(Balise{1});

Num_Fig = 0;

for k = [2:size(Data,2)] %On traite une ROI à la fois
    
    %On détecte les peaks, leur amplitude (proms), leur durée (width)
    %Le minpeakprom = 1; c'est le seuil de détection des peaks
    [pks, locs, width, proms] = findpeaks(Data{:,k},Data{:,1},'MinPeakProminence',0.2, 'MinPeakDistance',2.5);
    
    if isempty(pks) == 1
        pks = 0; locs = 0; width = 0; proms = 0;
    end
    
    Avant = find(locs<Balise); Apres = find(locs>Balise);
    
    Amp_mean_before = mean(proms(Avant)); Amp_mean_after = mean(proms(Apres));
    Dur_mean_before = mean(width(Avant)); Dur_mean_after = mean(width(Apres));
    Nb_before = size(pks(Avant),1); Nb_after = size(pks(Apres),1);
    
    Moyennes_before = [Nb_before, Amp_mean_before, Dur_mean_before]; Moyennes_before(isnan(Moyennes_before))=0;
    Moyennes_after = [Nb_after, Amp_mean_after, Dur_mean_after]; Moyennes_after(isnan(Moyennes_after))=0;
         
    Cellule = ['A',num2str(k-1)];
    Cellule_mean_before = ['A',num2str(k)];
    Cellule_mean_after = ['E',num2str(k)];
    
    Num_Fig = Num_Fig + 1;
    if Num_Fig == 10
        Num_Fig = 1;
    end
    
    figure(floor((k-2)/9)+1)
    subplot(3,3, Num_Fig)
    plot(Data{:,1},Data{:,k}, locs, pks, 'or')
    
    writematrix(Moyennes_before, 'Analyse_2P.xlsx', 'Sheet', 'Moyennes', 'Range', Cellule_mean_before)
    writematrix(Moyennes_after, 'Analyse_2P.xlsx', 'Sheet', 'Moyennes', 'Range', Cellule_mean_after)
    writematrix(proms', 'Analyse_2P.xlsx', 'Sheet', 'Amplitude', 'Range', Cellule)
    writematrix(width', 'Analyse_2P.xlsx', 'Sheet', 'Duration', 'Range', Cellule)
end