function output = summerinthecity

filepath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Raw Data\UK Solar Data.xlsx';
[~,~,r]=xlsread(filepath);

disp(strcat('Reading from file',32,filepath))

summer=struct();

devcount=0;

for a=2:2:length(r)
    if ~isempty(strfind(r{a},'Project developer'))
        devcount=devcount+1;
        if ~isempty(strfind(r{a}((strfind(r{a},'Project developer')+19):end),']'))
            dev=r{a}(strfind(r{a},'Project developer')+19:(strfind(r{a}((strfind(r{a},'Project developer')+19):end),']')+(strfind(r{a},'Project developer')+19)-1));
        elseif ~isempty(strfind(r{a}((strfind(r{a},'Project developer')+19):end),'Wiki'))
            dev=r{a}(strfind(r{a},'Project developer')+19:(strfind(r{a}((strfind(r{a},'Project developer')+19):end),'Wiki')+(strfind(r{a},'Project developer')+19)-2));
        elseif ~isempty(strfind(r{a}((strfind(r{a},'Project developer')+19):end),'In scheme'))
            dev=r{a}(strfind(r{a},'Project developer')+19:(strfind(r{a}((strfind(r{a},'Project developer')+19):end),'In scheme')+(strfind(r{a},'Project developer')+19)-2));
        end
        devstamp=matlab.lang.makeValidName(dev,'ReplacementStyle','delete');
    else
        dev='Unknown Developer';
        devstamp='UnknownDeveloper';
    end

    if ~isfield(summer,devstamp)
        summer.(devstamp)=[];
        disp(strcat('Found developer:',32,dev))
    end
    
    devlist{a/2}=dev;

    project=matlab.lang.makeValidName(r{a-1},'ReplacementStyle','delete');
    projectname=(r{a-1});
    
    spc=strfind(r{a},32);
    location={r{a}(1:spc(length(strfind(r{a},','))+1)-1)};
    
    if ~isempty(strfind((r{a}),'MWp'))
        spc=strfind(r{a}(1:strfind(r{a},'MWp')-2),32);
        MWp=str2double(r{a}( (spc(end)+1): (strfind(r{a},'MWp')-2) ));
    else
        MWp=[];
    end
    
    if ~isempty(strfind((r{a}),'MWAC'))
        spc=strfind(r{a}(1:strfind(r{a},'MWAC')-2),32);
        MWAC=str2double(r{a}( (spc(end)+1): (strfind(r{a},'MWAC')-2) ));
    else
        MWAC=[];
    end
    
    if ~isempty(strfind((r{a}),'Building'))
        status='Building';
    elseif ~isempty(strfind((r{a}),'Operating'))
        status='Operating';
    elseif ~isempty(strfind((r{a}),'Planned'))
        status='Planned';
    else
        status='Unknown';
    end
    
    if ~isempty(strfind((r{a}),'Area'))
        area=r{a}((strfind(r{a},'Area')+6):(strfind(r{a},' ha'))-1);
    else
        area='Unknown';
    end
    
    if ~isempty(strfind((r{a}),'owner'))
        owner=r{a}((strfind(r{a},'owner')+13):(strfind(r{a},'['))-2);
    else
        owner='Unknown';
    end
    
    if ~isempty(strfind((r{a}),'value'))
        value=r{a}((strfind(r{a},'value')+10):(strfind(r{a},'Wiki'))-2);
    else
        value='Unknown';
    end
    
    source='http://www.wiki-solar.org/maps/rg/W-S_MA~Rg%3Du-uN-GB.kml';
    
    summer.(devstamp).(project)=struct('Developer',dev,'Project',projectname,'Location',location','MWp',MWp,...
        'MWAC',MWAC,'Status',status,'Area',area,'Owner',owner,'Value',value,'Source',source);
end

devcount=length(fieldnames(summer));
devs=fieldnames(summer);

output=cell(750,9);

row=1;

for a=1:1:devcount
    row=row+1;
    projects=fieldnames(summer.(devs{a}));
    output{row,1}=summer.(devs{a}).(projects{1}).Developer;
    row=row+1;
    
    output{row,1}='Project Name';
    output{row,2}='MWp';
    output{row,3}='MWAC';
    output{row,4}='Status';
    output{row,5}='Location';
    output{row,6}='Area';
    output{row,7}='Plant Owner';
    output{row,8}='Project Value';
    output{row,9}='Source';
    row=row+1;
    
    for b=1:1:length(projects)
        output{row,1}=summer.(devs{a}).(projects{b}).Project;
        output{row,2}=summer.(devs{a}).(projects{b}).MWp;
        output{row,3}=summer.(devs{a}).(projects{b}).MWAC;
        output{row,4}=summer.(devs{a}).(projects{b}).Status;
        output{row,5}=summer.(devs{a}).(projects{b}).Location;
        output{row,6}=summer.(devs{a}).(projects{b}).Area;
        output{row,7}=summer.(devs{a}).(projects{b}).Owner;
        output{row,8}=summer.(devs{a}).(projects{b}).Value;
        output{row,9}=summer.(devs{a}).(projects{b}).Source;
        row=row+1;
    end
end

pivotoutput=cell(500,10);

row=1;
    
pivotoutput{row,1}='Developer Name';
pivotoutput{row,2}='Project Name';
pivotoutput{row,3}='MWp';
pivotoutput{row,4}='MWAC';
pivotoutput{row,5}='Status';
pivotoutput{row,6}='Location';
pivotoutput{row,7}='Area';
pivotoutput{row,8}='Plant Owner';
pivotoutput{row,9}='Project Value';
pivotoutput{row,10}='Source';
row=row+1;

for a=1:1:devcount 
    projects=fieldnames(summer.(devs{a}));   
    for b=1:1:length(projects)
        pivotoutput{row,1}=summer.(devs{a}).(projects{b}).Developer;
        pivotoutput{row,2}=summer.(devs{a}).(projects{b}).Project;
        pivotoutput{row,3}=summer.(devs{a}).(projects{b}).MWp;
        pivotoutput{row,4}=summer.(devs{a}).(projects{b}).MWAC;
        pivotoutput{row,5}=summer.(devs{a}).(projects{b}).Status;
        pivotoutput{row,6}=summer.(devs{a}).(projects{b}).Location;
        pivotoutput{row,7}=summer.(devs{a}).(projects{b}).Area;
        pivotoutput{row,8}=summer.(devs{a}).(projects{b}).Owner;
        pivotoutput{row,9}=summer.(devs{a}).(projects{b}).Value;
        pivotoutput{row,10}=summer.(devs{a}).(projects{b}).Source;
        row=row+1;
    end
end

%% Initialisation of POI Libs
% Add Java POI Libs to matlab javapath
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-3.8-20120326.jar');
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-ooxml-3.8-20120326.jar');
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-ooxml-schemas-3.8-20120326.jar');
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\xmlbeans-2.3.0.jar');
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\dom4j-1.6.1.jar');
javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\stax-api-1.0.1.jar');

outpath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Results\Results.xlsx';
xlwrite(outpath,output)

outpath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Results\Pivot Results.xlsx';
xlwrite(outpath,pivotoutput)