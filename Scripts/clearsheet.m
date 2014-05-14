function clearsheet (filename,sheetname)

% Retrieve sheet names 
[~, sheetNames] = xlsfinfo(filename);
cellsheetname={sheetname};
if ~isempty(strfind(sheetNames,sheetname))
    % Open Excel as a COM Automation server
    Excel = actxserver('Excel.Application');
    % Open Excel workbook
    Workbook = Excel.Workbooks.Open(filename);
    % Clear the content of the sheets (from the second onwards)
    cellfun(@(x) Excel.ActiveWorkBook.Sheets.Item(x).Cells.Clear, cellsheetname);
    % Now save/close/quit/delete
    Workbook.Save;
    Excel.Workbook.Close;
    invoke(Excel, 'Quit');
    delete(Excel)
else
    disp('Specified sheet does not exist!')
end