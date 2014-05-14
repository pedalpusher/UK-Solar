function [output,pivotoutput] = butterflystroke (write)

filepath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Raw Data\W-S_MA~Rg=u-uN-GB.kml';
fileID=fopen(filepath);

larva=textscan(fileID,'%s','delimiter',{'\n'});

soul=larva{1};

disp(strcat('Reading from file',32,filepath))
disp(strcat('File contains',32,(num2str(length(soul))),32,'lines'))

butterfly=struct();

devcount=0;

for a=1:1:length(soul)
    if strcmp(soul{a}(1:5),'<Plac')
        if ~isempty(strfind(soul{a},'Project developer'))
            devcount=devcount+1;
            spc=(strfind(soul{a},'Project developer'))+26;
            dev=soul{a}( spc(1) : (spc(1) + (strfind(soul{a}(spc(1):end),'<br>')-2) ) );
            devstamp=matlab.lang.makeValidName(dev,'ReplacementStyle','delete');
        else
            dev='Unknown Developer';
            devstamp='UnknownDeveloper';
        end

        if ~isfield(butterfly,devstamp)
            butterfly.(devstamp)=[];
            disp(strcat('Found developer:',32,dev))
        end

        projectname=soul{a}( (strfind(soul{a},'<name>'))+6:(strfind(soul{a},'</name>'))-1 );
        project=matlab.lang.makeValidName(projectname,'ReplacementStyle','delete');

        location={soul{a}((strfind(soul{a},'table width'))+26: ( ((strfind(soul{a},'table width'))+26) + strfind(soul{a}((strfind(soul{a},'table width'))+25:end),'<br>')-3) )};

        if ~isempty(strfind((soul{a}),'MWp'))
            MWp=soul{a}( (strfind(soul{a},'orange')+8):(strfind(soul{a},'MWp')-6) );
        else
            MWp=[];
        end

        if ~isempty(strfind((soul{a}),'<sub>AC'))
            MWAC=soul{a}( (strfind(soul{a},'red')+5):(strfind(soul{a},'<sub>AC')-8) );
        else
            MWAC=[];
        end

        if ~isempty(strfind((soul{a}),'Building'))
            status='Under construction';
        elseif ~isempty(strfind((soul{a}),'Operating'))
            status='Operating';
        elseif ~isempty(strfind((soul{a}),'Planned'))
            status='Planned';
        else
            status='Unknown';
        end

        if ~isempty(strfind((soul{a}),'Area'))
            area=soul{a}( (strfind(soul{a},'Area')+13):(strfind(soul{a},' ha')-1) );
        else
            area='Unknown';
        end

        if ~isempty(strfind((soul{a}),'Plant owner / IPP'))
            spc=(strfind(soul{a},'Plant owner / IPP'))+26;
            owner=soul{a}(spc(1): ( (spc(1)) + strfind(soul{a}(spc(1):end),'<br>')-2) );
        else
            owner='Unknown';
        end

        if ~isempty(strfind((soul{a}),'value'))
            value=soul{a}((strfind(soul{a},'value:'))+16: ( ((strfind(soul{a},'value:'))+16) + strfind(soul{a}((strfind(soul{a},'value:'))+16:end),'<')-2) );
        else
            value='Unknown';
        end

        if ~isempty(strfind((soul{a}),'land-owner'))
            spc=(strfind(soul{a},'land-owner'))+19;
            landowner=soul{a}(spc(1): ( spc(1) + strfind(soul{a}(spc(1):end),'<br>')-2) );
        else
            landowner='Unknown';
        end

        source='http://www.wiki-solar.org/maps/rg/W-S_MA~Rg%3Du-uN-GB.kml';

        butterfly.(devstamp).(project)=struct('Developer',dev,'Project',projectname,'Location',location','MWp',MWp,...
            'MWAC',MWAC,'Status',status,'Area',area,'ProjectOwner',owner,'LandOwner',landowner,'Value',value,'Source',source);
    end
end

disp('All lines checked and read')
disp('Writing results to format...')

devcount=length(fieldnames(butterfly));
devs=fieldnames(butterfly);

output=cell(750,10);

row=1;

for a=1:1:devcount
    row=row+1;
    projects=fieldnames(butterfly.(devs{a}));
    output{row,1}=butterfly.(devs{a}).(projects{1}).Developer;
    row=row+1;
    
    output{row,1}='Project Name';
    output{row,2}='MWp';
    output{row,3}='MWAC';
    output{row,4}='Status';
    output{row,5}='Location';
    output{row,6}='Area';
    output{row,7}='Plant Owner';
    output{row,8}='Land Owner';
    output{row,9}='Project Value';
    output{row,10}='Source';
    row=row+1;
    
    for b=1:1:length(projects)
        output{row,1}=butterfly.(devs{a}).(projects{b}).Project;
        output{row,2}=butterfly.(devs{a}).(projects{b}).MWp;
        output{row,3}=butterfly.(devs{a}).(projects{b}).MWAC;
        output{row,4}=butterfly.(devs{a}).(projects{b}).Status;
        output{row,5}=butterfly.(devs{a}).(projects{b}).Location;
        output{row,6}=butterfly.(devs{a}).(projects{b}).Area;
        output{row,7}=butterfly.(devs{a}).(projects{b}).ProjectOwner;
        output{row,8}=butterfly.(devs{a}).(projects{b}).LandOwner;
        output{row,9}=butterfly.(devs{a}).(projects{b}).Value;
        output{row,10}=butterfly.(devs{a}).(projects{b}).Source;
        row=row+1;
    end
end

pivotoutput=cell(500,11);

row=1;
    
pivotoutput{row,1}='Developer Name';
pivotoutput{row,2}='Project Name';
pivotoutput{row,3}='MWp';
pivotoutput{row,4}='MWAC';
pivotoutput{row,5}='Status';
pivotoutput{row,6}='Location';
pivotoutput{row,7}='Area (ha)';
pivotoutput{row,8}='Plant Owner';
pivotoutput{row,9}='Land Owner';
pivotoutput{row,10}='Project Value (£mn)';
pivotoutput{row,11}='Source';
row=row+1;

for a=1:1:devcount 
    projects=fieldnames(butterfly.(devs{a}));   
    for b=1:1:length(projects)
        pivotoutput{row,1}=butterfly.(devs{a}).(projects{b}).Developer;
        pivotoutput{row,2}=butterfly.(devs{a}).(projects{b}).Project;
        pivotoutput{row,3}=butterfly.(devs{a}).(projects{b}).MWp;
        pivotoutput{row,4}=butterfly.(devs{a}).(projects{b}).MWAC;
        pivotoutput{row,5}=butterfly.(devs{a}).(projects{b}).Status;
        pivotoutput{row,6}=butterfly.(devs{a}).(projects{b}).Location;
        pivotoutput{row,7}=butterfly.(devs{a}).(projects{b}).Area;
        pivotoutput{row,8}=butterfly.(devs{a}).(projects{b}).ProjectOwner;
        pivotoutput{row,9}=butterfly.(devs{a}).(projects{b}).LandOwner;
        pivotoutput{row,10}=butterfly.(devs{a}).(projects{b}).Value;
        pivotoutput{row,11}=butterfly.(devs{a}).(projects{b}).Source;
        row=row+1;
    end
end

disp('...done!')

if write==true
    disp('Writing results to excel...')
    
    %% Initialisation of POI Libs
    % Add Java POI Libs to matlab javapath
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-3.8-20120326.jar');
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-ooxml-3.8-20120326.jar');
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\poi-ooxml-schemas-3.8-20120326.jar');
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\xmlbeans-2.3.0.jar');
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\dom4j-1.6.1.jar');
    javaaddpath('C:\Users\davidchristopherson\Documents\MATLAB\~Switching Behavior Modelling\Scripts\Imported\poi_library\stax-api-1.0.1.jar');

    outpath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Results\Results.xlsx';
    worksheet='Results';
    EraseExcelSheets(outpath,worksheet);
    xlwrite(outpath,output,workesheet);

    outpath='C:\Users\davidchristopherson\Documents\MATLAB\~UK Solar\Results\Pivot Results.xlsx';
    EraseExcelSheets(outpath);
    xlwrite(outpath,pivotoutput);
    
    disp('...done!')
end