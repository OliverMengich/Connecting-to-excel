

%%
NET.addAssembly('microsoft.office.interop.excel');
app = Microsoft.Office.Interop.Excel.ApplicationClass;
books = app.Workbooks;
newWB = Add(books);
app.Visible = true;
%%
sheets = newWB.Worksheets;
newSheet = Item(sheets,1);
%%
newWS = Microsoft.Office.Interop.Excel.Worksheet(newSheet);
%%
excelArray = rand(10);

newRange = Range(newWS,'A1');
newRange.Value2 = 'Oliver';

newRange = Range(newWS,'A2:B11');
newRange.Value2 = excelArray;

newRange = Range(newWS,'B1');
newRange.Value2 = 'Kipkemei';


%%
strArray = NET.createArray('System.Object',3,1);
%%
strArray(1,1) = 'Add';
strArray(2,1) = 'text';
strArray(3,1) = 'to column C';
newRange = Range(newWS,'C3:C5');
newRange.Value2 = strArray;

%%
newFont =  newRange.Font;
newFont.Bold = 1;
newWS.Name = 'Test Data';

%%
SaveAs(newWB,'Example.csv');

Close(newWB)
Quit(app)
