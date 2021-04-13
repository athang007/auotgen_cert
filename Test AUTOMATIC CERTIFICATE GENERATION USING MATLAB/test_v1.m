clc
close all
home

filename = 'List.xls';
[num,txt] = xlsread(filename);

len = length(txt);

I = imread('certi.tif');

for i=1:len
    for j = 2:2
        text_name(i,j) = txt(i,j);
    end
end

for i=1:len
    for j = 3:3
        text_team(i,j) = txt(i,j);
    end
end

for i=2:len
    text_all = [text_name(i,2) text_team(i,3)]
    
   % position = [750 790; 1050 900];
   
    position = [ 1209.5 850 ; 1209.5 950 ]
    
    RGB = insertText(I,position,text_all,'FontSize',75,'BoxOpacity',0, 'AnchorPoint', 'Center', 'Font', 'VNlucida Handwriting'); %Font depends on what fonts you have
    
    figure
    
    imshow(RGB)
    
    y=i-1
        filename=['Result' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
end