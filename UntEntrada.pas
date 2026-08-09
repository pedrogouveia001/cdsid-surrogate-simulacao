unit UntEntrada;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompButton,
  IWCompListbox, IWCompEdit, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWCompLabel, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, IWGrids, ComObj, Dialogs,
  Variants, Math, UserSessionUnit, IWExtCtrls, jpeg, windows, ActiveX,
  IWWebGrid, IWAdvWebGrid, IWCompTabControl, IWAdvWebGridExcel, IWCompText,
  IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent, IWBaseHTML40Component,
  chart, series, pngimage, Grids,
  IWCompProgressBar, ExtCtrls, IWHTMLControls, TeeGDIPlus, TeEngine, TeeProcs,
  IWTMSPopup, IWHelpTip, Menus,

  BubbleCh, TeePolar,
  IWCompRadioButton, IWTMSCheckList, TeeLisB,
  TeeSeriesTextEd, Buttons, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Messages, Graphics, ComCtrls, StdCtrls, ErrorBar,
  IWProgressIndicator, IWDBGrids,
  IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, GanttCh;

type
  TIWForm1 = class(TIWAppForm)
    rgnmatcons: TIWRegion;
    IWLabel49: TIWLabel;
    IWImage31: TIWImage;
    IWImage32: TIWImage;
    IWLabel20: TIWLabel;
    IWRegion1: TIWRegion;
    IWLabel3: TIWLabel;
    IWLabel1: TIWLabel;
    cmbbxcriteria: TIWComboBox;
    IWLabel2: TIWLabel;
    btnshowresults: TIWButton;
    rgnresultado: TIWRegion;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWLabel7: TIWLabel;
    chrtResultado: TChart;
    rgnprimeiro: TIWRegion;
    grdresultados: TIWGrid;
    lbltexto: TIWLabel;
    imgResultado: TIWImage;
    Chart1: TChart;
    Series2: TBarSeries;
    edtnumalt: TIWEdit;
    grdImport: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWRegion2: TIWRegion;
    grdnsolutions: TIWGrid;
    IWLabel4: TIWLabel;
    IWLink1: TIWLink;
    IWLink2: TIWLink;
    IWLink3: TIWLink;
    rgnestatistica: TIWRegion;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWLabel9: TIWLabel;
    IWLabel10: TIWLabel;
    lblmedia: TIWLabel;
    lblmediana: TIWLabel;
    lblmoda: TIWLabel;
    lbldesviopadrao: TIWLabel;
    IWLabel8: TIWLabel;
    chcklstbxMetodos: TTIWCheckListBox;
    IWRegion5: TIWRegion;
    IWImage3: TIWImage;
    IWRegion7: TIWRegion;
    Series1: TBarSeries;
    IWImage4: TIWImage;
    IWRegion3: TIWRegion;
    IWRegion4: TIWRegion;
    rgnresultadoPromethee: TIWRegion;
    IWImage5: TIWImage;
    IWImage6: TIWImage;
    IWLabel11: TIWLabel;
    IWLabel12: TIWLabel;
    imgresultadopromethee: TIWImage;
    IWLabel13: TIWLabel;
    IWLink4: TIWLink;
    IWLink5: TIWLink;
    IWImage8: TIWImage;
    rgnestatisticapromethee: TIWRegion;
    IWLabel14: TIWLabel;
    IWLabel15: TIWLabel;
    IWLabel16: TIWLabel;
    IWLabel17: TIWLabel;
    lblmediapromethee: TIWLabel;
    lblmedianapromethee: TIWLabel;
    lblmodapromethee: TIWLabel;
    lbldesviopadraopromethee: TIWLabel;
    chrtresultadopromethee: TChart;
    BarSeries1: TBarSeries;
    IWRegion10: TIWRegion;
    grdnsolutionsPromethee: TIWGrid;
    IWRegion11: TIWRegion;
    IWRegion12: TIWRegion;
    rgnregradecisao: TIWRegion;
    grdregradecisao: TIWGrid;
    IWLink7: TIWLink;
    rgnregradecisaoPromethee: TIWRegion;
    grdregradecisaopromethee: TIWGrid;
    IWLink6: TIWLink;
    IWLabel41: TIWLabel;
    IWLabel18: TIWLabel;
    IWLabel19: TIWLabel;
    IWLabel21: TIWLabel;
    rgnAnaliseEstatistica: TIWRegion;
    IWImage12: TIWImage;
    IWImage13: TIWImage;
    IWLabel26: TIWLabel;
    IWLink12: TIWLink;
    IWText2: TIWText;
    lnkExportar: TIWLink;
    IWLabel63: TIWLabel;
    IWRegion18: TIWRegion;
    grdestatistica: TIWGrid;
    IWRegion19: TIWRegion;
    grdmaxmin: TIWGrid;
    IWImage14: TIWImage;
    IWRegion9: TIWRegion;
    IWRegion17: TIWRegion;
    IWLabel22: TIWLabel;
    IWLabel23: TIWLabel;
    IWText1: TIWText;
    IWText3: TIWText;
    rgnresultadosprobabilidade: TIWRegion;
    IWImage7: TIWImage;
    IWImage9: TIWImage;
    IWLabel24: TIWLabel;
    IWLabel25: TIWLabel;
    imgresultadosprobabilidade: TIWImage;
    IWLink8: TIWLink;
    IWImage11: TIWImage;
    IWLabel28: TIWLabel;
    IWRegion14: TIWRegion;
    IWRegion15: TIWRegion;
    IWRegion16: TIWRegion;
    grdresultadosprobabilidade: TIWGrid;
    rgnregra: TIWRegion;
    IWImage10: TIWImage;
    IWImage15: TIWImage;
    IWLabel27: TIWLabel;
    IWLabel29: TIWLabel;
    imgregra: TIWImage;
    IWLink9: TIWLink;
    IWImage17: TIWImage;
    IWLabel30: TIWLabel;
    IWRegion8: TIWRegion;
    IWRegion13: TIWRegion;
    IWRegion20: TIWRegion;
    grdregra: TIWGrid;
    chrtregra: TChart;
    BarSeries2: TBarSeries;
    IWRegion6: TIWRegion;
    IWLabel39: TIWLabel;
    cmbbxprobalt: TIWComboBox;
    IWLabel32: TIWLabel;
    chrtprobalt1: TChart;
    Series4: TBarSeries;
    chrtprobalt2: TChart;
    BarSeries3: TBarSeries;
    chrtprobalt3: TChart;
    BarSeries4: TBarSeries;
    grdmediaprobalt: TIWGrid;
    rgnalternativaX: TIWRegion;
    IWImage16: TIWImage;
    IWImage18: TIWImage;
    IWLabel31: TIWLabel;
    IWLabel33: TIWLabel;
    imgalternativaX: TIWImage;
    IWLink10: TIWLink;
    IWImage20: TIWImage;
    IWLabel34: TIWLabel;
    IWRegion22: TIWRegion;
    IWRegion23: TIWRegion;
    IWRegion24: TIWRegion;
    grdalternativaX: TIWGrid;
    chrtalternativaX: TChart;
    IWLink11: TIWLink;
    BarSeries7: TLineSeries;
    Series3: TLineSeries;
    Series5: TLineSeries;
    Series6: TLineSeries;
    rgnalternativaZ: TIWRegion;
    IWImage19: TIWImage;
    IWImage21: TIWImage;
    IWLabel35: TIWLabel;
    IWLabel36: TIWLabel;
    imgAlternativaZ: TIWImage;
    IWLink14: TIWLink;
    IWImage23: TIWImage;
    IWLabel37: TIWLabel;
    IWRegion25: TIWRegion;
    IWRegion26: TIWRegion;
    IWRegion27: TIWRegion;
    grdalternativaZ: TIWGrid;
    chrtalternativaZ: TChart;
    LineSeries1: TLineSeries;
    LineSeries2: TLineSeries;
    LineSeries3: TLineSeries;
    LineSeries4: TLineSeries;
    chrtalternativaX2: TChart;
    LineSeries5: TLineSeries;
    LineSeries6: TLineSeries;
    LineSeries7: TLineSeries;
    LineSeries8: TLineSeries;
    chrtAlternativaZ2: TChart;
    LineSeries9: TLineSeries;
    LineSeries10: TLineSeries;
    LineSeries11: TLineSeries;
    LineSeries12: TLineSeries;
    Series7: TLineSeries;
    Series8: TLineSeries;
    Series9: TLineSeries;
    Series10: TLineSeries;
    rgnalternativaY: TIWRegion;
    IWImage22: TIWImage;
    IWImage24: TIWImage;
    IWLabel38: TIWLabel;
    IWLabel40: TIWLabel;
    imgalternativaY: TIWImage;
    IWLink15: TIWLink;
    IWImage26: TIWImage;
    IWLabel42: TIWLabel;
    IWRegion28: TIWRegion;
    IWRegion29: TIWRegion;
    IWRegion30: TIWRegion;
    grdalternativaY: TIWGrid;
    chrtalternativaY: TChart;
    LineSeries14: TLineSeries;
    LineSeries15: TLineSeries;
    lblproblema1: TIWLabel;
    lblproblema2: TIWLabel;
    lblproblema3: TIWLabel;
    lblproblema4: TIWLabel;
    lblproblema5: TIWLabel;
    rgnalternativaw: TIWRegion;
    IWImage25: TIWImage;
    IWImage27: TIWImage;
    IWLabel43: TIWLabel;
    IWLabel44: TIWLabel;
    imgalternativaw: TIWImage;
    IWLink17: TIWLink;
    IWImage29: TIWImage;
    IWLabel45: TIWLabel;
    lblproblema6: TIWLabel;
    IWRegion31: TIWRegion;
    IWRegion32: TIWRegion;
    IWRegion33: TIWRegion;
    grdalternativaw: TIWGrid;
    chrtalternativaw: TChart;
    LineSeries13: TLineSeries;
    LineSeries16: TLineSeries;
    IWLink18: TIWLink;
    rgnalternativaS: TIWRegion;
    IWImage28: TIWImage;
    IWImage30: TIWImage;
    IWLabel46: TIWLabel;
    IWLabel47: TIWLabel;
    imgalternativaS: TIWImage;
    IWLink19: TIWLink;
    IWImage34: TIWImage;
    IWLabel48: TIWLabel;
    lblproblema7: TIWLabel;
    IWRegion34: TIWRegion;
    IWRegion35: TIWRegion;
    IWRegion36: TIWRegion;
    grdalternativaS: TIWGrid;
    chrtalternativaS: TChart;
    LineSeries17: TLineSeries;
    LineSeries18: TLineSeries;
    LineSeries19: TLineSeries;
    LineSeries20: TLineSeries;
    LineSeries21: TLineSeries;
    lblPercentualSigualaX: TIWLabel;
    IWButton1: TIWButton;
    IWButton2: TIWButton;
    IWButton3: TIWButton;
    IWButton4: TIWButton;
    rgnperdamaxima: TIWRegion;
    IWImage33: TIWImage;
    IWImage35: TIWImage;
    IWLabel50: TIWLabel;
    IWLabel51: TIWLabel;
    imgperdamaxima: TIWImage;
    IWLink13: TIWLink;
    IWImage37: TIWImage;
    IWLabel52: TIWLabel;
    lblproblema8: TIWLabel;
    IWRegion37: TIWRegion;
    IWRegion38: TIWRegion;
    IWRegion39: TIWRegion;
    grdperdamaxima: TIWGrid;
    chrtperdamaxima: TChart;
    LineSeries22: TLineSeries;
    LineSeries23: TLineSeries;
    LineSeries24: TLineSeries;
    IWLink16: TIWLink;
    rgnConsequenceRange: TIWRegion;
    IWImage36: TIWImage;
    IWImage38: TIWImage;
    IWLabel53: TIWLabel;
    IWLabel54: TIWLabel;
    imgConsequenceRange: TIWImage;
    IWLink20: TIWLink;
    IWImage40: TIWImage;
    IWLabel55: TIWLabel;
    lblproblema9: TIWLabel;
    IWRegion40: TIWRegion;
    IWRegion41: TIWRegion;
    IWRegion42: TIWRegion;
    grdConsequenceRange: TIWGrid;
    chrtConsequenceRange: TChart;
    LineSeries25: TLineSeries;
    LineSeries27: TLineSeries;
    IWLink21: TIWLink;
    rgndifvg: TIWRegion;
    IWImage39: TIWImage;
    IWImage41: TIWImage;
    IWLabel56: TIWLabel;
    IWLabel57: TIWLabel;
    imgDifVG: TIWImage;
    IWLink22: TIWLink;
    IWImage43: TIWImage;
    IWLabel58: TIWLabel;
    lblproblema10: TIWLabel;
    IWRegion43: TIWRegion;
    IWRegion44: TIWRegion;
    IWRegion45: TIWRegion;
    grdDifVG: TIWGrid;
    chrtDifVG: TChart;
    LineSeries29: TLineSeries;
    LineSeries30: TLineSeries;
    IWLink23: TIWLink;
    rgnmatrizconseqAltXS: TIWRegion;
    IWImage42: TIWImage;
    IWImage44: TIWImage;
    IWLabel59: TIWLabel;
    imgmatrizconseqAltXS: TIWImage;
    IWLink24: TIWLink;
    IWImage46: TIWImage;
    IWLabel61: TIWLabel;
    IWLabel62: TIWLabel;
    IWRegion46: TIWRegion;
    IWRegion47: TIWRegion;
    chrtpercentilMatrizConseqXS: TChart;
    IWLink25: TIWLink;
    cmbbxmatrizconseqAltXS: TIWComboBox;
    chrtmatrizconseqAltXS: TChart;
    LineSeries31: TLineSeries;
    LineSeries32: TLineSeries;
    Series11: TLineSeries;
    Series12: TLineSeries;
    Series13: TLineSeries;
    Series14: TLineSeries;
    Series15: TLineSeries;
    Series16: TLineSeries;
    Series17: TLineSeries;
    Series18: TLineSeries;
    Series19: TLineSeries;
    Series20: TLineSeries;
    Series21: TLineSeries;
    Series22: TLineSeries;
    Series23: TLineSeries;
    Series24: TLineSeries;
    Series25: TLineSeries;
    Series26: TLineSeries;
    Series27: TLineSeries;
    Series28: TLineSeries;
    LineSeries26: TBarSeries;
    LineSeries28: TBarSeries;
    grdexportarresultados: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO2: TTIWAdvWebGridExcelIO;
    lnkMatrizConseqAltX: TIWLink;
    grdMatrizConseqAltX: TTIWAdvWebGrid;
    procedure cmbbxauxiliarAsyncChange(Sender: TObject;
      EventParams: TStringList);
    procedure cmbbxauxiliarChange(Sender: TObject);
    procedure btnshowresultsClick(Sender: TObject);
    procedure GerarCases;
    procedure Surrogate;
    procedure rgnresultadoCreate(Sender: TObject);
    procedure SetarCores;
    procedure rgnmatconsCreate(Sender: TObject);
    function Mean(const Data: array of Double): Double;
    function StandardDeviation(const Data: array of Double; IsSample: Boolean)
      : Double;
    procedure IWLink1Click(Sender: TObject);
    procedure LimpaVariaveis;
    procedure IWLink2Click(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure ComparacaoParAPar;
    procedure Promethee;
    // procedure ResultadosPromethee;
    procedure CalculoFluxos;
    procedure rgnresultadoPrometheeCreate(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink6Click(Sender: TObject);
    procedure IWLink12Click(Sender: TObject);
    procedure rgnresultadosprobabilidadeCreate(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure rgnregraCreate(Sender: TObject);
    procedure IWLink9Click(Sender: TObject);
    procedure cmbbxprobaltChange(Sender: TObject);
    procedure IWLink11Click(Sender: TObject);
    procedure IWLink10Click(Sender: TObject);
    procedure OrdenarColunas;
    procedure QuartisPorColuna;
    procedure IWLink13Click(Sender: TObject);
    procedure IWLink14Click(Sender: TObject);
    procedure rgnalternativaZCreate(Sender: TObject);
    procedure rgnalternativaYCreate(Sender: TObject);
    procedure IWLink16Click(Sender: TObject);
    procedure IWLink15Click(Sender: TObject);
    procedure IWLink17Click(Sender: TObject);
    procedure IWLink18Click(Sender: TObject);
    procedure rgnalternativawCreate(Sender: TObject);
    procedure rgnalternativaSCreate(Sender: TObject);
    procedure IWLink19Click(Sender: TObject);
    procedure IWButton1Click(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure IWButton3Click(Sender: TObject);
    procedure IWButton4Click(Sender: TObject);
    procedure rgnperdamaximaCreate(Sender: TObject);
    procedure IWLink21Click(Sender: TObject);
    procedure IWLink20Click(Sender: TObject);
    procedure rgnConsequenceRangeCreate(Sender: TObject);
    procedure IWLink23Click(Sender: TObject);
    procedure IWLink22Click(Sender: TObject);
    procedure rgndifvgCreate(Sender: TObject);
    procedure IWLink25Click(Sender: TObject);
    procedure IWLink24Click(Sender: TObject);
    procedure rgnmatrizconseqAltXSCreate(Sender: TObject);
    procedure cmbbxmatrizconseqAltXSChange(Sender: TObject);
    procedure ExportarResultados;
    procedure lnkMatrizConseqAltXClick(Sender: TObject);
    // procedure EscalaRazao;
  public
  end;

implementation

{$R *.dfm}

function TIWForm1.Mean(const Data: array of Double): Double;
var
  Sum: Double;
  I: Integer;
begin
  Sum := 0;
  for I := Low(Data) to High(Data) do
    Sum := Sum + Data[I];
  Result := Sum / Length(Data);
end;

function TIWForm1.StandardDeviation(const Data: array of Double;
  IsSample: Boolean): Double;
var
  MeanValue, Sum: Double;
  I: Integer;
  N: Integer;
begin
  if Length(Data) = 0 then
    raise Exception.Create('Data array is empty.');

  MeanValue := Mean(Data);
  Sum := 0;

  for I := Low(Data) to High(Data) do
    Sum := Sum + Sqr(Data[I] - MeanValue);

  N := Length(Data);
  if IsSample then
    Dec(N); // For sample standard deviation, use n-1 in the denominator

  Result := Sqrt(Sum / N);
end;


procedure TIWForm1.QuartisPorColuna;
  // calcula Q1,Q2,Q3 de cada coluna da matriz Mat
  // saída Q: numCols linhas x 3 colunas (Q[c,0]=Q1, Q[c,1]=Q2, Q[c,2]=Q3)
var
  numRows, numCols: Integer;


  // retorna percentil p (0..1) de um vetor JÁ ORDENADO (interpolação linear)
  function PercentilInterpolado(const SortedArr: array of Double; p: Double): Double;
  var
    n: Integer;
    h: Double;
    k: Integer;
    d: Double;
  begin
    Result := 0;
    n := Length(SortedArr);
    if n = 0 then Exit;

    if p <= 0 then
    begin
      Result := SortedArr[0];
      Exit;
    end;
    if p >= 1 then
    begin
      Result := SortedArr[n - 1];
      Exit;
    end;

    h := p * (n + 1);

    if h <= 1 then
    begin
      Result := SortedArr[0];
      Exit;
    end
    else if h >= n then
    begin
      Result := SortedArr[n - 1];
      Exit;
    end;

    k := Trunc(h); // floor(h), 1-based
    d := h - k;    // fractional part

    // indices 0-based: k -> SortedArr[k-1], k+1 -> SortedArr[k]
    Result := SortedArr[k - 1] + d * (SortedArr[k] - SortedArr[k - 1]);
  end;

var
  c, r,i: Integer;
  aux : double;
  Col: array of Double;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
  // para a alternativa X
  numRows := Length(VetorAlternativaX);
  if numRows = 0 then
  begin
    SetLength(QuartisAlternativaX, 0, 0);
    Exit;
  end;

  numCols := Length(VetorAlternativaX[0]);

  // checagem de retangularidade
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaX[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // aloca saída: numCols linhas x 3 colunas


  for c := 0 to numCols - 1 do
  begin
    // copia coluna
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).VetorAlternativaX[r][c];

    // ordena e calcula quartis

    QuartisAlternativaX[c][0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisAlternativaX[c][1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisAlternativaX[c][2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisAlternativaX[c][3] := PercentilInterpolado(Col, 1.00); // Q4

    for i := 0 to 100 do
    begin
      aux := i /100;
      TIWUserSession(WebApplication.Data).PercentilAlternativaX[c][i] := PercentilInterpolado(Col, aux);
    end;
  end;

  // Para a alternativa Z
  numRows := Length(VetorAlternativaZ);
  if numRows = 0 then
  begin
    SetLength(VetorAlternativaZ, 0, 0);
    Exit;
  end;

  numCols := Length(VetorAlternativaZ[0]);

  // checagem de retangularidade
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaZ[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // aloca saída: numCols linhas x 3 colunas


  for c := 0 to numCols - 1 do
  begin
    // copia coluna
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).VetorAlternativaZ[r][c];

    // ordena e calcula quartis

    QuartisAlternativaZ[c][0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisAlternativaZ[c][1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisAlternativaZ[c][2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisAlternativaZ[c][3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilAlternativaZ[c][i] := PercentilInterpolado(Col, aux);
    end;
  end;



  // Para a alternativa S
  numRows := Length(VetorAlternativaS);
  if numRows = 0 then
  begin
    SetLength(VetorAlternativaS, 0, 0);
    Exit;
  end;

  numCols := Length(VetorAlternativaS[0]);

  // checagem de retangularidade
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaS[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // aloca saída: numCols linhas x 3 colunas


  for c := 0 to numCols - 1 do
  begin
    // copia coluna
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).VetorAlternativaS[r][c];

    // ordena e calcula quartis

    TIWUserSession(WebApplication.Data).QuartisAlternativaS[c][0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisAlternativaS[c][1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisAlternativaS[c][2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisAlternativaS[c][3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilAlternativaS[c][i] := PercentilInterpolado(Col, aux);
    end;
  end;


  // Para a alternativa y
  numRows := Length(VetoralternativaY);
  if numRows = 0 then
  begin
    SetLength(VetoralternativaY, 0, 0);
    Exit;
  end;

  numCols := Length(VetoralternativaY[0]);

  // checagem de retangularidade
  for r := 0 to numRows - 1 do
    if Length(VetoralternativaY[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // aloca saída: numCols linhas x 3 colunas


  for c := 0 to numCols - 1 do
  begin
    // copia coluna
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).VetoralternativaY[r][c];

    // ordena e calcula quartis

    QuartisalternativaY[c][0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisalternativaY[c][1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisalternativaY[c][2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisalternativaY[c][3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilalternativaY[c][i] := PercentilInterpolado(Col, aux);
    end;
   end;




  // Para a alternativa w
  numRows := Length(VetorAlternativaw);
  if numRows = 0 then
  begin
    SetLength(VetorAlternativaw, 0, 0);
    Exit;
  end;

  numCols := Length(VetorAlternativaW[0]);

  // checagem de retangularidade
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaW[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // aloca saída: numCols linhas x 3 colunas


  for c := 0 to numCols - 1 do
  begin
    // copia coluna
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).VetorAlternativaW[r][c];

    // ordena e calcula quartis

    QuartisAlternativaW[c][0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisAlternativaW[c][1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisAlternativaW[c][2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisAlternativaW[c][3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilAlternativaW[c][i] := PercentilInterpolado(Col, aux);
    end;
  end;
   // Uma vez que obtivemos os percentis das alternativas de interesse pelas suas probabilidades de
  // ocupar uma determinada posição, vamos fazer o mesmo para suas perdas maximas

  // Alternativa X
  numRows := Length(PerdaMaximaX);
  if numRows = 0 then
  begin
    SetLength(PerdaMaximaX, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).PerdaMaximaX[r];

    // ordena e calcula quartis

    QuartisPerdaMaximaX[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisPerdaMaximaX[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisPerdaMaximaX[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisPerdaMaximaX[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilPerdaMaximaX[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa Z
  numRows := Length(PerdaMaximaZ);
  if numRows = 0 then
  begin
    SetLength(PerdaMaximaZ, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).PerdaMaximaZ[r];

    // ordena e calcula quartis

    QuartisPerdaMaximaZ[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisPerdaMaximaZ[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisPerdaMaximaZ[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisPerdaMaximaZ[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilPerdaMaximaZ[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa S
  numRows := Length(PerdaMaximaS);
  if numRows = 0 then
  begin
    SetLength(PerdaMaximaS, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).PerdaMaximaS[r];

    // ordena e calcula quartis

    QuartisPerdaMaximaS[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisPerdaMaximaS[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisPerdaMaximaS[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisPerdaMaximaS[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilPerdaMaximaS[i] := PercentilInterpolado(Col, aux);
    end;


   // Analisando agora a maxima distancia entre as consequencias para a X, Z e S
    // Alternativa X
  numRows := Length(ConsequenceRangeX);
  if numRows = 0 then
  begin
    SetLength(ConsequenceRangeX, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).ConsequenceRangeX[r];

    // ordena e calcula quartis

    QuartisConsequenceRangeX[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisConsequenceRangeX[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisConsequenceRangeX[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisConsequenceRangeX[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilConsequenceRangeX[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa Z
  numRows := Length(ConsequenceRangeZ);
  if numRows = 0 then
  begin
    SetLength(ConsequenceRangeZ, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).ConsequenceRangeZ[r];

    // ordena e calcula quartis

    QuartisConsequenceRangeZ[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisConsequenceRangeZ[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisConsequenceRangeZ[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisConsequenceRangeZ[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilConsequenceRangeZ[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa S
  numRows := Length(ConsequenceRangeS);
  if numRows = 0 then
  begin
    SetLength(ConsequenceRangeS, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).ConsequenceRangeS[r];

    // ordena e calcula quartis

    QuartisConsequenceRangeS[0] := PercentilInterpolado(Col, 0.25); // Q1
    QuartisConsequenceRangeS[1] := PercentilInterpolado(Col, 0.50); // Q2
    QuartisConsequenceRangeS[2] := PercentilInterpolado(Col, 0.75); // Q3
    QuartisConsequenceRangeS[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).PercentilConsequenceRangeS[i] := PercentilInterpolado(Col, aux);
    end;


    // Analisando agora a maxima distancia entre vg para a X, Z e S
    // Alternativa X
  numRows := Length(difVGX);
  if numRows = 0 then
  begin
    SetLength(difVGX, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).difVGX[r];

    // ordena e calcula quartis

    quartisdifVGX[0] := PercentilInterpolado(Col, 0.25); // Q1
    quartisdifVGX[1] := PercentilInterpolado(Col, 0.50); // Q2
    quartisdifVGX[2] := PercentilInterpolado(Col, 0.75); // Q3
    quartisdifVGX[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).percentildifVGX[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa Z
  numRows := Length(difVGZ);
  if numRows = 0 then
  begin
    SetLength(difVGZ, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).difVGZ[r];

    // ordena e calcula quartis

    quartisDifVGZ[0] := PercentilInterpolado(Col, 0.25); // Q1
    quartisDifVGZ[1] := PercentilInterpolado(Col, 0.50); // Q2
    quartisDifVGZ[2] := PercentilInterpolado(Col, 0.75); // Q3
    quartisDifVGZ[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).percentildifVGZ[i] := PercentilInterpolado(Col, aux);
    end;

  // Alternativa S
  numRows := Length(difVGS);
  if numRows = 0 then
  begin
    SetLength(difVGS, 0);
    Exit;
  end;

    // copia as linhas
    SetLength(Col, numRows);
    for r := 0 to numRows - 1 do
      Col[r] := TIWUserSession(WebApplication.Data).difVGS[r];

    // ordena e calcula quartis

    quartisdifVGS[0] := PercentilInterpolado(Col, 0.25); // Q1
    quartisdifVGS[1] := PercentilInterpolado(Col, 0.50); // Q2
    quartisdifVGS[2] := PercentilInterpolado(Col, 0.75); // Q3
    quartisdifVGS[3] := PercentilInterpolado(Col, 1.00); // Q4


    for i := 0 to 100 do
    begin
      aux := i/100;
      TIWUserSession(WebApplication.Data).percentildifVGS[i] := PercentilInterpolado(Col, aux);
    end;


  end;
end;


procedure TIWForm1.OrdenarColunas;
var
  numRows, numCols: Integer;
  r, c, i, j, minIdx, minIdx2: Integer;
  temp: Double;

  temp2 : array of double;
begin
   with TIWUserSession(WebApplication.Data) do
  begin

  // Para a alternativa X
  // verifica se matriz vazia
  if Length(VetorAlternativaX) = 0 then Exit;
  numRows := Length(VetorAlternativaX);
  numCols := Length(VetorAlternativaX[0]);

  // checar retangularidade (opcional, mas recomendado)
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaX[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // para cada coluna, aplicar selection sort nas linhas
  for c := 0 to numCols - 1 do
  begin
    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if VetorAlternativaX[j][c] < VetorAlternativaX[minIdx][c] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := VetorAlternativaX[i][c];
        VetorAlternativaX[i][c] := VetorAlternativaX[minIdx][c];
        TIWUserSession(WebApplication.Data).VetorAlternativaX[minIdx][c] := temp;
      end;
    end;
  end;

  // Para a alternativa Z
  // verifica se matriz vazia
  if Length(VetorAlternativaZ) = 0 then Exit;
  numRows := Length(VetorAlternativaZ);
  numCols := Length(VetorAlternativaZ[0]);

  // checar retangularidade (opcional, mas recomendado)
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaZ[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // para cada coluna, aplicar selection sort nas linhas
  for c := 0 to numCols - 1 do
  begin
    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if VetorAlternativaZ[j][c] < VetorAlternativaZ[minIdx][c] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := VetorAlternativaZ[i][c];
        VetorAlternativaZ[i][c] := VetorAlternativaZ[minIdx][c];
        TIWUserSession(WebApplication.Data).VetorAlternativaZ[minIdx][c] := temp;
      end;
    end;
  end;


  // Para a alternativa S
  // verifica se matriz vazia
  if Length(VetorAlternativaZ) = 0 then Exit;
  numRows := Length(VetorAlternativaZ);
  numCols := Length(VetorAlternativaZ[0]);

  // checar retangularidade (opcional, mas recomendado)
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaS[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // para cada coluna, aplicar selection sort nas linhas
  for c := 0 to numCols - 1 do
  begin
    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if VetorAlternativaS[j][c] < VetorAlternativaS[minIdx][c] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := VetorAlternativaS[i][c];
        VetorAlternativaS[i][c] := VetorAlternativaS[minIdx][c];
        TIWUserSession(WebApplication.Data).VetorAlternativaS[minIdx][c] := temp;
      end;
    end;
  end;


  // Para a alternativa Y
  // verifica se matriz vazia
  if Length(VetoralternativaY) = 0 then Exit;
  numRows := Length(VetoralternativaY);
  numCols := Length(VetoralternativaY[0]);

  // checar retangularidade (opcional, mas recomendado)
  for r := 0 to numRows - 1 do
    if Length(VetoralternativaY[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // para cada coluna, aplicar selection sort nas linhas
  for c := 0 to numCols - 1 do
  begin
    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if VetoralternativaY[j][c] < VetoralternativaY[minIdx][c] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := VetoralternativaY[i][c];
        VetoralternativaY[i][c] := VetoralternativaY[minIdx][c];
        TIWUserSession(WebApplication.Data).VetoralternativaY[minIdx][c] := temp;
      end;
    end;
  end;


  // Para a alternativa w
  // verifica se matriz vazia
  if Length(VetorAlternativaW) = 0 then Exit;
  numRows := Length(VetorAlternativaW);
  numCols := Length(VetorAlternativaW[0]);

  // checar retangularidade (opcional, mas recomendado)
  for r := 0 to numRows - 1 do
    if Length(VetorAlternativaW[r]) <> numCols then
      raise Exception.Create('Matriz não é retangular (linhas com tamanhos diferentes).');

  // para cada coluna, aplicar selection sort nas linhas
  for c := 0 to numCols - 1 do
  begin
    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if VetorAlternativaW[j][c] < VetorAlternativaW[minIdx][c] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := VetorAlternativaW[i][c];
        VetorAlternativaW[i][c] := VetorAlternativaW[minIdx][c];
        TIWUserSession(WebApplication.Data).VetorAlternativaW[minIdx][c] := temp;
      end;
    end;
  end;


  // Uma vez que ordenamos as alternativas de interesse pelas suas probabilidades de
  // ocupar uma determinada posição, vamos fazer o mesmo para suas perdas maximas

  // Alternativa X
  // verifica se matriz vazia
  finalize(temp2);
  setlength(temp2, numcrit);
    if Length(PerdaMaximaX) = 0 then Exit;
    numRows := Length(PerdaMaximaX);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if PerdaMaximaX[j] < PerdaMaximaX[minIdx] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := PerdaMaximaX[i];
        PerdaMaximaX[i] := PerdaMaximaX[minIdx];
        TIWUserSession(WebApplication.Data).PerdaMaximaX[minIdx] := temp;

        j := ordemPerdaMaximaX[i];
        ordemPerdaMaximaX[i] := ordemPerdaMaximaX[minIdx];
        ordemPerdaMaximaX[minIdx] := j;


        for c := 0 to numalt - 1 do
        for j := 0 to numcrit - 1 do
        begin
          temp := MatrizConseqPercentilPerdaX[i,c,j];
          MatrizConseqPercentilPerdaX[i,c,j] := MatrizConseqPercentilPerdaX[minIdx,c,j];
          MatrizConseqPercentilPerdaX[minIdx,c,j] := temp;
        end;


        for j := 0 to numcrit - 1 do
        begin
          temp2[j] := matrizconseqaltx[i,j];
          matrizconseqaltx[i,j] := matrizconseqaltx[minIdx,j];
          matrizconseqaltx[minIdx,j] := temp2[j];
        end;

        for j := 0 to numcrit - 1 do
        begin
          temp2[j] := matrizconseqaltS[i,j];
          matrizconseqaltS[i,j] := matrizconseqaltS[minIdx,j];
          matrizconseqaltS[minIdx,j] := temp2[j];
        end;

        for j := 0 to numcrit - 1 do
        begin
          temp2[j] := matrizconseqaltxOrdenado[i,j];
          matrizconseqaltxOrdenado[i,j] := matrizconseqaltxOrdenado[minIdx,j];
          matrizconseqaltxOrdenado[minIdx,j] := temp2[j];
        end;

        for j := 0 to numcrit - 1 do
        begin
          temp2[j] := matrizconseqaltSOrdenado[i,j];
          matrizconseqaltSOrdenado[i,j] := matrizconseqaltSOrdenado[minIdx,j];
          matrizconseqaltSOrdenado[minIdx,j] := temp2[j];
        end;
      end;
    end;

    for I := 0 to ninst - 1 do
    begin
      for j := 0 to numcrit - 1 do
      for r := 0 to numcrit - 1 do
      begin
        if j = r then continue
        else
        begin
          if matrizconseqaltxOrdenado[i,j] > matrizconseqaltxOrdenado[i,r] then
          begin
            temp := matrizconseqaltxOrdenado[i,j];
            matrizconseqaltxOrdenado[i,j] := matrizconseqaltxOrdenado[i,r];
            matrizconseqaltxOrdenado[i,r] := temp;
          end;

          if matrizconseqaltSOrdenado[i,j] > matrizconseqaltSOrdenado[i,r] then
          begin
            temp := matrizconseqaltSOrdenado[i,j];
            matrizconseqaltSOrdenado[i,j] := matrizconseqaltSOrdenado[i,r];
            matrizconseqaltSOrdenado[i,r] := temp;
          end;

        end;

      end;
    end;

  // Alternativa Z
  if Length(PerdaMaximaZ) = 0 then Exit;
    numRows := Length(PerdaMaximaZ);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if PerdaMaximaZ[j] < PerdaMaximaZ[minIdx] then
          minIdx := j;
      end;
      if minIdx <> i then
      begin
        temp := PerdaMaximaZ[i];
        PerdaMaximaZ[i] := PerdaMaximaZ[minIdx];
        TIWUserSession(WebApplication.Data).PerdaMaximaZ[minIdx] := temp;


        j := ordemPerdaMaximaZ[i];
        ordemPerdaMaximaZ[i] := ordemPerdaMaximaZ[minIdx];
        ordemPerdaMaximaZ[minIdx] := j;
      end;
    end;

  // Alternativa S
  finalize(temp2);
  setlength(temp2, numcrit);
  if Length(PerdaMaximaS) = 0 then Exit;
    numRows := Length(PerdaMaximaS);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if PerdaMaximaS[j] < PerdaMaximaS[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := PerdaMaximaS[i];
        PerdaMaximaS[i] := PerdaMaximaS[minIdx];
        TIWUserSession(WebApplication.Data).PerdaMaximaS[minIdx] := temp;


        j := ordemPerdaMaximaS[i];
        ordemPerdaMaximaS[i] := ordemPerdaMaximaS[minIdx];
        ordemPerdaMaximaS[minIdx] := j;
      end; // if




    end; // for i


    // Analisando a diferença maxima entre as consequencias para X, S e Z
    //Para X
    If Length(ConsequenceRangeX) = 0 then Exit;
    numRows := Length(ConsequenceRangeX);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if ConsequenceRangeX[j] < ConsequenceRangeX[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := ConsequenceRangeX[i];
        ConsequenceRangeX[i] := ConsequenceRangeX[minIdx];
        TIWUserSession(WebApplication.Data).ConsequenceRangeX[minIdx] := temp;
      end; // if
    end; // for i

    //Para Z
    If Length(ConsequenceRangez) = 0 then Exit;
    numRows := Length(ConsequenceRangez);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if ConsequenceRangez[j] < ConsequenceRangez[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := ConsequenceRangez[i];
        ConsequenceRangez[i] := ConsequenceRangez[minIdx];
        TIWUserSession(WebApplication.Data).ConsequenceRangez[minIdx] := temp;
      end; // if
    end; // for i

    // Para S
    If Length(ConsequenceRangeS) = 0 then Exit;
    numRows := Length(ConsequenceRangeS);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if ConsequenceRangeS[j] < ConsequenceRangeS[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := ConsequenceRangeS[i];
        ConsequenceRangeS[i] := ConsequenceRangeS[minIdx];
        TIWUserSession(WebApplication.Data).ConsequenceRangeS[minIdx] := temp;
      end; // if
    end; // for i


    // Analisando a diferença de vg entre as consequencias para X, S e Z
    //Para X
    If Length(difVGX) = 0 then Exit;
    numRows := Length(difVGX);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if difVGX[j] < difVGX[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := difVGX[i];
        difVGX[i] := difVGX[minIdx];
        TIWUserSession(WebApplication.Data).difVGX[minIdx] := temp;
      end; // if
    end; // for i

    //Para Z
    If Length(difVGZ) = 0 then Exit;
    numRows := Length(difVGZ);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if difVGZ[j] < difVGZ[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := difVGZ[i];
        difVGZ[i] := difVGZ[minIdx];
        TIWUserSession(WebApplication.Data).difVGZ[minIdx] := temp;
      end; // if
    end; // for i

    // Para S
    If Length(difVGS) = 0 then Exit;
    numRows := Length(difVGS);
    //numCols := Length(VetorAlternativaW[0]);

    for i := 0 to numRows - 2 do
    begin
      minIdx := i;
      for j := i + 1 to numRows - 1 do
      begin
        if difVGS[j] < difVGS[minIdx] then
          minIdx := j;
      end; // for j
      if minIdx <> i then
      begin
        temp := difVGS[i];
        difVGS[i] := difVGS[minIdx];
        TIWUserSession(WebApplication.Data).difVGS[minIdx] := temp;
      end; // if
    end; // for i

  end; // with
end;


{ // Transforma os valores da consMatrix em uma escala de razão
  procedure TIWForm1.EscalaRazao;
  var
  i, j: integer;
  Maxl, Minl, MaxA, MinA, MaxB, MinB: Array of Double;
  AuxVetPOA: array of integer;
  // ConsRazao: array of array of Double;
  begin

  with TIWUserSession(WebApplication.Data) do
  begin
  Setlength(Consrazao, NumAlt, numcrit);
  Setlength(Maxl, numcrit);
  Setlength(Minl, numcrit);
  Setlength(AuxVetPOA, Length(VtrAlternatives));

  Setlength(MaxA, numcrit);
  Setlength(MinA, numcrit);
  Setlength(MaxB, numcrit);
  Setlength(MinB, numcrit);

  for i := 0 to Length(VtrAlternatives) - 1 do
  begin
  AuxVetPOA[i] := TIWUserSession(WebApplication.Data).VtrAlternatives[i];
  end;

  if TIWUserSession(WebApplication.Data).Problematica <> 3 then
  begin
  // ** Buscando maximos e minimos locais (Escala de Razão) **//
  for j := 0 to numcrit - 1 do
  begin
  Maxl[j] := ConsMatrix[AuxVetPOA[0], j];
  Minl[j] := ConsMatrix[AuxVetPOA[0], j];
  end;
  for j := 0 to numcrit - 1 do
  for i := 0 to Length(VtrAlternatives) - 1 do
  begin
  if ConsMatrix[AuxVetPOA[i], j] > Maxl[j] then
  Maxl[j] := ConsMatrix[AuxVetPOA[i], j];
  if ConsMatrix[AuxVetPOA[i], j] < Minl[j] then
  Minl[j] := ConsMatrix[AuxVetPOA[i], j];
  end;

  for j := 0 to numcrit - 1 do
  for i := 0 to Length(VtrAlternatives) - 1 do
  begin
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) or (Typeofcrit[j] = 0)
  or (Typeofcrit[j] = 2) then
  begin // dos criterios de maximização

  // * ConsequenciaPositivas ou zero *//
  if (Maxl[j] > 0) and (Minl[j] >= 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  (TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j] / Maxl[j]);
  // fim caso para consequencias são diferente de zero

  // * Consequencias mistas (+, -, 0)*//
  if (Maxl[j] >= 0) and (Minl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j]
  / (Maxl[j] - Minl[j]) + 1 - Maxl[j] / (Maxl[j] - Minl[j]);

  // * Consequencias negativas *//
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  Maxl[j] / TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j];

  end; // Fimdoscritérios de maximização

  { if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  begin // inicio dos critérios de minimização

  // * ConsequenciaPositivas *//
  if (ConsMatrix[i, j] <> 0) and (Minl[j] > 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  (Minl[j] / TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j]);
  // para consequencias diferentes de zero

  // * Consequencias mistas (+, -, 0)*//
  if (Minl[j] <= 0) and (Maxl[j] >= 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j] /
  (Minl[j] - Maxl[j]) + 1 - Minl[j] / (Minl[j] - Maxl[j]);

  // * Consequencias negativas *//
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j]
  / Minl[j];

  end; // fim dos critérios de minimização
  end; // i
  end;
  if TIWUserSession(WebApplication.Data).Problematica = 3 then
  begin
  for j := 0 to numcrit - 1 do
  begin
  if (AuxVetPOA[0] >= 0) and (AuxVetPOA[1] < 0) then
  begin
  MaxA[j] := ConsMatrix[AuxVetPOA[0], j];
  MinA[j] := ConsMatrix[AuxVetPOA[0], j];
  MaxB[j] := TIWUserSession(WebApplication.Data)
  .altperfil[-AuxVetPOA[1] - 1, j];
  MinB[j] := altperfil[-AuxVetPOA[1] - 1, j];
  end;

  if (AuxVetPOA[0] < 0) and (AuxVetPOA[1] >= 0) then
  begin
  MaxA[j] := ConsMatrix[AuxVetPOA[1], j];
  MinA[j] := ConsMatrix[AuxVetPOA[1], j];
  MaxB[j] := TIWUserSession(WebApplication.Data)
  .altperfil[-AuxVetPOA[0] - 1, j];
  MinB[j] := altperfil[-AuxVetPOA[0] - 1, j];
  end;

  if (AuxVetPOA[0] >= 0) and (AuxVetPOA[1] >= 0) then
  begin
  MaxA[j] := ConsMatrix[AuxVetPOA[0], j];
  MinA[j] := ConsMatrix[AuxVetPOA[0], j];
  MaxB[j] := ConsMatrix[AuxVetPOA[0], j];
  MinB[j] := ConsMatrix[AuxVetPOA[0], j];
  end;
  end;

  for j := 0 to numcrit - 1 do
  begin
  if MaxA[j] >= MaxB[j] then
  begin
  Maxl[j] := MaxA[j];
  end
  else
  begin
  Maxl[j] := MaxB[j];
  end;
  if MinA[j] <= MinB[j] then
  begin
  Minl[j] := MinA[j];
  end
  else
  Minl[j] := MinB[j];
  end;

  for j := 0 to numcrit - 1 do
  for i := 0 to Length(VtrAlternatives) - 1 do
  begin
  if AuxVetPOA[i] >= 0 then
  begin
  if TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j]
  > Maxl[j] then
  Maxl[j] := ConsMatrix[AuxVetPOA[i], j];
  if ConsMatrix[AuxVetPOA[i], j] < Minl[j] then
  Minl[j] := ConsMatrix[AuxVetPOA[i], j];
  end
  else
  begin
  if TIWUserSession(WebApplication.Data)
  .altperfil[-AuxVetPOA[i] - 1, j] > Maxl[j] then
  Maxl[j] := altperfil[AuxVetPOA[i] * (-1) - 1, j];
  if altperfil[AuxVetPOA[i] * (-1) - 1, j] < Minl[j] then
  Minl[j] := altperfil[AuxVetPOA[i] * (-1) - 1, j];
  end;
  end;

  for j := 0 to numcrit - 1 do
  for i := 0 to Length(VtrAlternatives) - 1 do
  begin
  if AuxVetPOA[i] >= 0 then
  begin

  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) or
  (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) or
  (Typeofcrit[j] = 4) or (Typeofcrit[j] = 5) then
  begin // dos criterios de maximização

  // * ConsequenciaPositivas ou zero *//
  if (Maxl[j] > 0) and (Minl[j] >= 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j]
  := (TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j] / Maxl[j]);

  // fim caso para consequencias são diferente de zero

  // * Consequencias mistas (+, -, 0)*//
  if (Maxl[j] >= 0) and (Minl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j]
  := TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j] / (Maxl[j] - Minl[j])
  + 1 - Maxl[j] / (Maxl[j] - Minl[j]);

  // * Consequencias negativas *//
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j]
  := Maxl[j] / TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j];

  end; // Fimdoscritérios de maximização

  { if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  begin // inicio dos critérios de minimização

  // * ConsequenciaPositivas *//
  if (ConsMatrix[i, j] <> 0) and (Minl[j] > 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  (Minl[j] / TIWUserSession(WebApplication.Data)
  .ConsMatrix[AuxVetPOA[i], j]);
  // para consequencias diferentes de zero

  // * Consequencias mistas (+, -, 0)*//
  if (Minl[j] <= 0) and (Maxl[j] >= 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j] /
  (Minl[j] - Maxl[j]) + 1 - Minl[j] / (Minl[j] - Maxl[j]);

  // * Consequencias negativas *//
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  TIWUserSession(WebApplication.Data).Consrazao[AuxVetPOA[i], j] :=
  TIWUserSession(WebApplication.Data).ConsMatrix[AuxVetPOA[i], j]
  / Minl[j];

  end; // fim dos critérios de minimização
  end; // if
  end; // i

  // ****** ALTERNATIVAS PERFIS
  for j := 0 to numcrit - 1 do
  for i := 0 to Length(indice_altp) - 1 do
  begin
  if AuxVetPOA[i] < 0 then
  begin

  if Maxl[j] = Minl[j] then
  begin
  TIWUserSession(WebApplication.Data).Consrazaop[i, j] := 1;
  end
  else
  begin
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) then
  // dos criterios de maximização
  begin
  if (Maxl[j] > 0) and (Minl[j] >= 0) then
  // * ConsequenciaPositivas *//
  TIWUserSession(WebApplication.Data).Consrazaop[i, j] :=
  altperfil[indice_altp[i] - 1, j] / Maxl[j];
  if (Maxl[j] >= 0) and (Minl[j] < 0) then
  // * Consequencias mistas (+, -, 0)*//
  Consrazaop[i, j] := altperfil[indice_altp[i] - 1, j] /
  (Maxl[j] - Minl[j]) + 1 - Maxl[j] / (Maxl[j] - Minl[j]);
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  // * Consequencias negativas *//
  Consrazaop[i, j] := Maxl[j] / altperfil
  [indice_altp[i] - 1, j];
  end // Fimdoscritérios de maximização
  else // (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then // inicio dos critérios de minimização
  begin
  if (ConsMatrix[i, j] <> 0) and (Minl[j] > 0) then
  // * ConsequenciaPositivas *//
  Consrazaop[i, j] :=
  (Minl[j] / altperfil[indice_altp[i] - 1, j]);
  if (Minl[j] <= 0) and (Maxl[j] >= 0) then
  // * Consequencias mistas (+, -, 0)*//
  Consrazaop[i, j] := altperfil[indice_altp[i] - 1, j] /
  (Minl[j] - Maxl[j]) + 1 - Minl[j] / (Minl[j] - Maxl[j]);
  if (Minl[j] < 0) and (Maxl[j] < 0) then
  // * Consequencias negativas *//
  Consrazaop[i, j] := altperfil[indice_altp[i] - 1, j] / Minl
  [j];
  end; // fim dos critérios de minimização
  end; // else
  end; // for i := 0 to length(indice_alt) - 1 do
  end; // for j := 0 to NumCrit - 1 do
  end;

  end; // with
  end; // fim do procedimento }

procedure TiwForm1.ExportarResultados;
var
i, j, k : integer;
begin
 grdexportarresultados.ClearCells;
 grdexportarresultados.Columns.Clear;

 for i := 0 to 10 - 1 do
 grdexportarresultados.Columns.Add;

 k := 1 + grdalternativaX.RowCount;
 grdexportarresultados.TotalRows := k;


 grdexportarresultados.Cells[0,0] := 'Alternative X Results';
 for i := 0 to grdalternativaX.RowCount - 1 do
 for j := 0 to grdalternativaX.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,i+1] := grdalternativaX.Cell[i,j].Text;
 end;


 k := k + 1;
 grdexportarresultados.TotalRows := k + 1 +  grdalternativaS.RowCount;


 grdexportarresultados.Cells[0,k] := 'Alternative S Results';
 for i := 0 to grdalternativaS.RowCount - 1 do
 for j := 0 to grdalternativaS.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,K+i+1] := grdalternativaS.Cell[i,j].Text;
 end;

 K := grdexportarresultados.TotalRows+1;
 grdexportarresultados.TotalRows := k + 1 + grdalternativaZ.RowCount;

 grdexportarresultados.Cells[0,k] := 'Alternative Z Results';
 for i := 0 to grdalternativaZ.RowCount - 1 do
 for j := 0 to grdalternativaZ.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,K+i+1] := grdalternativaZ.Cell[i,j].Text;
 end;

 K := grdexportarresultados.TotalRows+1;
 grdexportarresultados.TotalRows := k + 1 + grdalternativaY.RowCount;

 grdexportarresultados.Cells[0,k] := 'Alternative Y Results';
 for i := 0 to grdalternativaY.RowCount - 1 do
 for j := 0 to grdalternativaY.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,K+i+1] := grdalternativaY.Cell[i,j].Text;
 end;

 K := grdexportarresultados.TotalRows+1;
 grdexportarresultados.TotalRows := k + 1 + grdalternativaW.RowCount;

 grdexportarresultados.Cells[0,k] := 'Alternative W Results';
 for i := 0 to grdalternativaW.RowCount - 1 do
 for j := 0 to grdalternativaW.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,K+i+1] := grdalternativaW.Cell[i,j].Text;
 end;

 K := grdexportarresultados.TotalRows+1;
 grdexportarresultados.TotalRows := k + 1 + grdperdamaxima.RowCount;

 grdexportarresultados.Cells[0,k] := 'Maximum Loss Results';
 for i := 0 to grdperdamaxima.RowCount - 1 do
 for j := 0 to grdperdamaxima.ColumnCount - 1 do
 begin
 grdexportarresultados.Cells[j,K+i+1] := grdperdamaxima.Cell[i,j].Text;
 end;

 tiwadvwebgridexcelio2.AdvWebGrid := grdexportarresultados;

 TIWAdvWebGridExcelIO2.XLSExport
        (TiwUserSession(WebApplication.Data).Nomeproblema + '   - Results' + '.xls', WebApplication);



end;

procedure TIWForm1.Surrogate;
var
  I, j, k, t, s, q, aux: Integer;
  auxpeso, soma, auxresultado, aux2, aux3: real;
  NovaSol: Boolean;
  Fatorial: real;
  auxsol: Integer;
  primeira, segunda, terceira, altz, altw, alty, alts, altx: Integer;
  pular1, pular2, pular3: Boolean;
  minIdx: Integer;
  tmp: Double;
  tmp2: array of Double;
  erro : boolean;
  auxprimeira, auxsegunda, auxterceira: array of real;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    finalize(MatrizPOA);
    finalize(ResultadoROC);
    finalize(MatrizSol);
    finalize(ResultAlt);
    finalize(ResultSol);
    finalize(CaseSol);

    finalize(MatrizPOAPromethee);
    finalize(ResultadoPromethee);
    finalize(MatrizSolPromethee);
    finalize(ResultAltPromethee);
    finalize(ResultSolPromethee);
    finalize(CaseSolPromethee);
    finalize(perdamedia);
    finalize(perdamaxima);
    finalize(probalt);

    { SetLength(PesoCrit, numcrit);
      SetLength(ResultadoROC, Length(CasesOrdemcrit), Numalt);
      SetLength(MatrizPOA, Length(CasesOrdemcrit), Numalt);
      SetLength(CaseSol, Length(CasesOrdemcrit)); }
    SetLength(ResultadoROC, Length(CasesOrdemcrit), Numalt);
    SetLength(ResultadoPromethee, Length(CasesOrdemcrit), Numalt);
    SetLength(MatrizDifVG, Length(CasesOrdemcrit), Numalt);
    SetLength(MatrizPOA, Length(CasesOrdemcrit), Numalt);
    SetLength(MatrizDifVGPromethee, Length(CasesOrdemcrit), Numalt);
    SetLength(MatrizPOAPromethee, Length(CasesOrdemcrit), Numalt);
    SetLength(MatrizRankingAlt, Length(CasesOrdemcrit), Numalt);
    SetLength(RankingPrimeira, Numalt);
    SetLength(RankingSegunda, Numalt);
    SetLength(RankingTerceira, Numalt);
    SetLength(CaseSol, Length(CasesOrdemcrit));
    SetLength(CaseSolPromethee, Length(CasesOrdemcrit));
    SetLength(perdamaxima, Numalt);
    SetLength(perdamedia, Numalt);
    SetLength(probalt, Numalt);

    SetLength(perdamaximaPromethee, Numalt);
    SetLength(perdamediaPromethee, Numalt);
    SetLength(probaltPromethee, Numalt);

    SetLength(Primcrit, numcrit);
    SetLength(Seccrit, numcrit);
    SetLength(terccrit, numcrit);
    SetLength(quartcrit, numcrit);

    for j := 0 to Numalt - 1 do
    begin
      perdamedia[j] := 0;
      perdamaxima[j] := 0;
      probalt[j] := 0;
    end;

    for j := 0 to Numalt - 1 do
    begin
      perdamediaPromethee[j] := 0;
      perdamaximaPromethee[j] := 0;
      probaltPromethee[j] := 0;
    end;

    for k := 0 to Length(CasesOrdemcrit) - 2 do
    begin
      if aditivo then
      begin
        for j := 0 to Numalt - 1 do
          for I := 0 to numcrit - 1 do
          begin
            TiwUserSession(WebApplication.Data).ResultadoROC[k, j] :=
              ResultadoROC[k, j] + MatrizConseqNorm[j, I] * TiwUserSession
              (WebApplication.Data).PesoCrit[k, I];
          end;
        auxresultado := 0;
        for j := 0 to Numalt - 1 do
        begin
          if ResultadoROC[k, j] > auxresultado then
            auxresultado := ResultadoROC[k, j];
        end;
        for j := 0 to Numalt - 1 do
        begin
          TiwUserSession(WebApplication.Data).MatrizDifVG[k, j] :=
            auxresultado - ResultadoROC[k, j];
          if auxresultado = ResultadoROC[k, j] then
          begin
            TiwUserSession(WebApplication.Data).MatrizPOA[k, j] := 1;
            TiwUserSession(WebApplication.Data).probalt[j] := probalt[j] + 1;
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if perdamaxima[j] < MatrizDifVG[k, j] then
            perdamaxima[j] := MatrizDifVG[k, j];

          perdamedia[j] := perdamedia[j] + MatrizDifVG[k, j];
        end;

        for j := 0 to Numalt - 1 do
          MatrizRankingAlt[k, j] := 0;

        for I := 0 to Numalt - 2 do
          for j := I + 1 to Numalt - 1 do
          begin
            // compara apenas resultadoroc (não será alterado)
            if TiwUserSession(WebApplication.Data).ResultadoROC[k, I]
              < TiwUserSession(WebApplication.Data).ResultadoROC[k, j] then
              inc(TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, I]);

            if TiwUserSession(WebApplication.Data).ResultadoROC[k, I]
              >= TiwUserSession(WebApplication.Data).ResultadoROC[k, j] then
              inc(TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, j]);

          end;

      end;

      if sobreclassificacao then
      begin
        // Promethee
        ComparacaoParAPar;
        auxpromethee := k;
        CalculoFluxos;

        for j := 0 to Numalt - 1 do
        begin
          ResultadoPromethee[k, j] := NetFlow[j];
        end;
        auxresultado := 0;
        for j := 0 to Numalt - 1 do
        begin
          if ResultadoPromethee[k, j] > auxresultado then
            auxresultado := ResultadoPromethee[k, j];
        end; // for j
        for j := 0 to Numalt - 1 do
        begin
          TiwUserSession(WebApplication.Data).MatrizDifVGPromethee[k, j] :=
            auxresultado - ResultadoPromethee[k, j];
          if auxresultado = ResultadoPromethee[k, j] then
          begin
            TiwUserSession(WebApplication.Data).MatrizPOAPromethee[k, j] := 1;
            probaltPromethee[j] := probaltPromethee[j] + 1;
          end;  // if
        end; // for j
        for j := 0 to Numalt - 1 do
        begin
          if perdamaximaPromethee[j] < MatrizDifVGPromethee[k, j] then
            perdamaximaPromethee[j] := MatrizDifVGPromethee[k, j];
          perdamediaPromethee[j] := perdamediaPromethee[j] +
            MatrizDifVGPromethee[k, j];
        end;   // for j

      end; // if sobreclassificacao
    end; // for k


    // caso especial para os pesos todos iguais
    if aditivo then
    begin
      for j := 0 to Numalt - 1 do
        for I := 0 to numcrit - 1 do
        begin
          TiwUserSession(WebApplication.Data)
            .ResultadoROC[Length(CasesOrdemcrit) - 1, j] := ResultadoROC[k, j]
            + MatrizConseqNorm[j, I] * (1 / numcrit);
        end;
      k := Length(CasesOrdemcrit) - 1;
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if ResultadoROC[k, j] > auxresultado then
          auxresultado := ResultadoROC[k, j];

      end;
      for j := 0 to Numalt - 1 do
      begin
        TiwUserSession(WebApplication.Data).MatrizDifVG[k, j] :=
          auxresultado - ResultadoROC[k, j];
        if auxresultado = ResultadoROC[k, j] then
        begin
          TiwUserSession(WebApplication.Data).MatrizPOA[k, j] := 1;
          TiwUserSession(WebApplication.Data).probalt[j] := probalt[j] + 1;
        end;

      end;

      for j := 0 to Numalt - 1 do
        MatrizRankingAlt[k, j] := 0;

      for I := 0 to Numalt - 1 do
        for j := 0 to Numalt - 1 do
        begin
        if i = j then continue;

          // compara apenas resultadoroc (não será alterado)
          if TiwUserSession(WebApplication.Data).MatrizDifVG[k, I]
            > TiwUserSession(WebApplication.Data).MatrizDifVG[k, j] then
            inc(TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, I]);

         { if TiwUserSession(WebApplication.Data).ResultadoROC[k, I]
            >= TiwUserSession(WebApplication.Data).ResultadoROC[k, j] then
            inc(TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, j]); }

        end;

      for j := 0 to Numalt - 1 do
      begin
        if perdamaxima[j] < MatrizDifVG[k, j] then
          TiwUserSession(WebApplication.Data).perdamaxima[j] := MatrizDifVG
            [k, j];
        TiwUserSession(WebApplication.Data).perdamedia[j] := perdamedia[j]
          + MatrizDifVG[k, j];
      end;

    end;
    for j := 0 to Numalt - 1 do
    begin
      if probalt[j] < Length(CasesOrdemcrit) then
        perdamedia[j] := perdamedia[j] / (Length(CasesOrdemcrit) - probalt[j]);
      if probalt[j] = Length(CasesOrdemcrit) then
        perdamedia[j] := 0;
    end;

    auxresultado := 0;
    primeira := 0;
    for j := 0 to Numalt - 1 do
    begin
      if probalt[j] > auxresultado then
      begin
        auxresultado := probalt[j];
        primeira := j;
      end;
    end;
    auxresultado := 0;
    segunda := 0;
    for j := 0 to Numalt - 1 do
    begin
      if (probalt[j] > auxresultado) and (primeira <> j) then
      begin
        auxresultado := probalt[j];
        segunda := j;
      end;
    end;
    auxresultado := 0;
    terceira := 0;
    for j := 0 to Numalt - 1 do
    begin
      if (probalt[j] > auxresultado) and (primeira <> j) and (segunda <> j) then
      begin
        auxresultado := probalt[j];
        terceira := j;
      end;
    end;

    mediaprobalt := (probalt[primeira] / (Length(CasesOrdemcrit)))
      + mediaprobalt;

    if (probalt[primeira] / (Length(CasesOrdemcrit))) > maximoprobalt then
      maximoprobalt := (probalt[primeira] / (Length(CasesOrdemcrit)));

    TiwUserSession(WebApplication.Data).probaltestratificado
      [floor(100 * (TiwUserSession(WebApplication.Data).probalt[primeira] /
          (Length(CasesOrdemcrit))))] := 1 + probaltestratificado
      [floor(100 * (probalt[primeira] / (Length(CasesOrdemcrit))))];

    if probalt[primeira] < Length(CasesOrdemcrit) then
    begin
      TiwUserSession(WebApplication.Data).probaltestratificado2
        [floor(100 * ((TiwUserSession(WebApplication.Data).probalt[segunda]
              + probalt[primeira]) / (Length(CasesOrdemcrit))))] := 1 +
        probaltestratificado2
        [floor(100 * ((probalt[segunda] + probalt[primeira]) /
            (Length(CasesOrdemcrit))))];

      mediaprobalt2 := mediaprobalt2 + (probalt[primeira] + probalt[segunda])
        / Length(CasesOrdemcrit);
    end
    else // se probalt[primeira] = length(casesordemcrit), então probestratificado de 1 e 2 é igual a probestratificado de 1
    begin
      TiwUserSession(WebApplication.Data).probaltestratificado2[100] :=
        1 + probaltestratificado2[100];

      mediaprobalt2 := mediaprobalt2 + 1;

      // TiwUserSession(WebApplication.Data).probaltestratificado3
      // [floor(100 * (TiwUserSession(WebApplication.Data).probalt[primeira] /
      // (Length(CasesOrdemcrit))))] := 1 + probaltestratificado3
      // [floor(100 * (probalt[primeira] / (Length(CasesOrdemcrit))))];
    end;
    if (probalt[primeira] + probalt[segunda]) < Length(CasesOrdemcrit) then
    begin
      TiwUserSession(WebApplication.Data).probaltestratificado3
        [floor(100 * ((TiwUserSession(WebApplication.Data).probalt[terceira]
              + probalt[segunda] + probalt[primeira]) / (Length(CasesOrdemcrit))
          ))] := 1 + probaltestratificado3
        [floor(100 * ((probalt[terceira] + probalt[segunda] + probalt[primeira]
            ) / (Length(CasesOrdemcrit))))];

      mediaprobalt3 := mediaprobalt3 +
        (probalt[primeira] + probalt[segunda] + probalt[terceira]) / Length
        (CasesOrdemcrit);

    end
    else
    begin
      TiwUserSession(WebApplication.Data).probaltestratificado3[100] :=
        1 + probaltestratificado3[100];

      mediaprobalt3 := mediaprobalt3 + 1;

    end;

    matrizprobalt[instancia] := TiwUserSession(WebApplication.Data).probalt[primeira] / (Length(CasesOrdemcrit));

    SetLength(auxprimeira, Numalt);
    SetLength(auxsegunda, Numalt);
    SetLength(auxterceira, Numalt);

   vetoralternativaY[instancia, 0] := 0;
   vetoralternativaX[instancia, 0] := 0;
   vetoralternativaX[instancia, 1] := 0;
   vetoralternativaX[instancia, 2] := 0;
   vetoralternativaX[instancia, 3] := 0;
   vetoralternativaX[instancia, 4] := 0;
   alty := -1;
   altw := -1;
   alts := -1;
   altx := -1;


   //Para a alt X
   for j := 0 to Numalt - 1 do
   begin
   aux2 := 0;
    for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (TiwUserSession(WebApplication.Data).MatrizRankingAlt[k,j] = 0) then
      aux2 := aux2 + 1;
    end;// for k
    if VetoralternativaX[instancia, 0] < aux2  then
    begin
      VetoralternativaX[instancia, 0] := aux2;
      altX := j;
    end; // if
    end; // for j


    for k := 0 to length(casesordemcrit) - 1 do
    begin
     if MatrizRankingAlt[k,altX] = 1 then
     VetoralternativaX[instancia, 1] := VetoralternativaX[instancia, 1] + 1;

     if MatrizRankingAlt[k,altX] = 2 then
     VetoralternativaX[instancia, 2] := VetoralternativaX[instancia, 2] + 1;

     if (MatrizRankingAlt[k,altX] = 0) or (MatrizRankingAlt[k,altX] = 1) then
     VetoralternativaX[instancia, 3] := VetoralternativaX[instancia, 3] + 1;

     if (MatrizRankingAlt[k,altX] = 0) or (MatrizRankingAlt[k,altX] = 1) or (MatrizRankingAlt[k,altX] = 2) then
     VetoralternativaX[instancia, 4] := VetoralternativaX[instancia, 4] + 1;
    end;




   for j := 0 to Numalt - 1 do
   begin
   aux2 := 0;
    for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (TiwUserSession(WebApplication.Data).MatrizRankingAlt[k,j] = 1) then
      aux2 := aux2 + 1;
    end;// for k
    if VetoralternativaS[instancia, 1] < aux2  then
    begin
      VetoralternativaS[instancia, 1] := aux2;
      alts := j;
    end; // if
    end; // for j

    if alts = primeira then    //  S = X
    percentualSigualaX := percentualSigualaX + 1;
    //VetoralternativaS[instancia, 5] := VetoralternativaS[instancia, 5] + 1;

    for k := 0 to length(casesordemcrit) - 1 do
    begin
     if MatrizRankingAlt[k,altS] = 0 then
     VetoralternativaS[instancia, 0] := VetoralternativaS[instancia, 0] + 1;

     if MatrizRankingAlt[k,altS] = 2 then
     VetoralternativaS[instancia, 2] := VetoralternativaS[instancia, 2] + 1;

     if (MatrizRankingAlt[k,altS] = 0) or (MatrizRankingAlt[k,altS] = 1) then
     VetoralternativaS[instancia, 3] := VetoralternativaS[instancia, 3] + 1;

     if (MatrizRankingAlt[k,altS] = 0) or (MatrizRankingAlt[k,altS] = 1) or (MatrizRankingAlt[k,altS] = 2) then
     VetoralternativaS[instancia, 4] := VetoralternativaS[instancia, 4] + 1;
    end;


   for j := 0 to Numalt - 1 do
   begin
   aux2 := 0;
    if j = primeira then continue;

    for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (TiwUserSession(WebApplication.Data).MatrizRankingAlt[k,j] = 1) and (MatrizRankingAlt[k,primeira] = 0) then
      aux2 := aux2 + 1;
    end;
    if VetoralternativaY[instancia, 0] < aux2  then
    begin
    VetoralternativaY[instancia, 0] := aux2;
    alty := j;
    end;
   end;

   aux2 := 0;
   if (alty <> -1) then
   begin
   for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (MatrizRankingAlt[k,alty] = 1) then
      aux2 := aux2 + 1;
    end;
    if VetoralternativaY[instancia, 1] < aux2  then
    VetoralternativaY[instancia, 1] := aux2;
   end;

   for i := 0 to 4 do
   begin
   TiwUserSession(WebApplication.Data).vetoralternativaS[instancia, i] := vetoralternativaS[instancia, i] / length(casesordemcrit);
   TiwUserSession(WebApplication.Data).vetoralternativaX[instancia, i] := vetoralternativaX[instancia, i] / length(casesordemcrit);
   end;
   TiwUserSession(WebApplication.Data).vetoralternativaY[instancia, 1] := vetoralternativaY[instancia, 1] / length(casesordemcrit);
   TiwUserSession(WebApplication.Data).vetoralternativaY[instancia, 0] := vetoralternativaY[instancia, 0] / length(casesordemcrit);


   for j := 0 to Numalt - 1 do
   begin
   aux2 := 0;
    if (j = altx) or (j = alty) then continue;

    for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (TiwUserSession(WebApplication.Data).MatrizRankingAlt[k,j] = 1) then
      aux2 := aux2 + 1;
    end;
    if VetorAlternativaW[instancia, 1] < aux2  then
    begin
    VetorAlternativaW[instancia, 1] := aux2;
    altW := j;
    end;
   end;

   aux2 := 0;
   if altw <> -1 then
   begin
     for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      if (TiwUserSession(WebApplication.Data).MatrizRankingAlt[k,altw] = 1) and (MatrizRankingAlt[k,primeira] = 0) then
      aux2 := aux2 + 1;
    end;
    if VetorAlternativaW[instancia, 0] < aux2  then
    begin
    VetorAlternativaW[instancia, 0] := aux2;
    //alty := j;
    end;
   end;

   TiwUserSession(WebApplication.Data).vetoralternativaW[instancia, 1] := vetoralternativaW[instancia, 1] / length(casesordemcrit);
   TiwUserSession(WebApplication.Data).vetoralternativaW[instancia, 0] := vetoralternativaW[instancia, 0] / length(casesordemcrit);

   for j := 0 to numalt - 1 do
   begin
   for i := 0 to numcrit - 1 do
   begin
     MatrizConseqPercentilPerdaX[instancia,j,i] := matrizconseqnorm[j, i];
     if j = altx then
     MatrizConseqAltX[instancia, i] := matrizconseqnorm[altX, i];
     if j = alts then
     MatrizConseqAltS[instancia, i] := matrizconseqnorm[altS, i];

     MatrizConseqAltXOrdenado[instancia, i] := MatrizConseqAltX[instancia, i];
     MatrizConseqAltSOrdenado[instancia, i] := MatrizConseqAltS[instancia, i];
   end;
   end;

   (**
      Uma vez calculado a probabilidade associada às alternativas de interesse X, Y, Z, W e S
       vamos calcular a perda maxima das alternativas X, Z e S para quando uma delas é a solução do problema
      **)
      PerdaMaximaX[instancia] := 0;
      PerdaMaximaS[instancia] := 0;
      PerdaMaximaZ[instancia] := 0;
      PerdaMaximaW[instancia] := 0;

      for k := 0 to length(casesordemcrit) - 1 do
      begin
      {if TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, altx] <> 0 then // X n é sol
      begin
       if MatrizDifVG[k, altx] > PerdaMaximaX[instancia] then
        begin
        PerdaMaximaX[instancia] := MatrizDifVG[k, altx];
        OrdemPerdaMaximaX[instancia] := k;

        if k = length(casesordemcrit) - 1  then
        PercentualPerdaMaximaXPesosIguais := PercentualPerdaMaximaXPesosIguais + 1;

        end;
      end;

      if TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, segunda] <> 0 then // Z n é sol
      begin
       if MatrizDifVG[k, segunda] > PerdaMaximaZ[instancia] then
        begin
        PerdaMaximaZ[instancia] := MatrizDifVG[k, Segunda];
        OrdemPerdaMaximaZ[instancia] := k;


        if k = length(casesordemcrit) - 1  then
        PercentualPerdaMaximaZPesosIguais := PercentualPerdaMaximaZPesosIguais + 1;

        end;
      end;

      if TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, alts] <> 0 then // S n é sol
      begin
       if MatrizDifVG[k, altS] > PerdaMaximaS[instancia] then
        begin
        PerdaMaximaS[instancia] := MatrizDifVG[k, altS];
        OrdemPerdaMaximaS[instancia] := k;


        if k = length(casesordemcrit) - 1  then
        PercentualPerdaMaximaSPesosIguais := PercentualPerdaMaximaSPesosIguais + 1;
        end;
      end;  }



      // Quando X é sol, podemos calcular a perda maxima para S e Z e assim sucessivamente
        if TiwUserSession(WebApplication.Data).MatrizRankingAlt[k, altx] = 0 then // X é sol
        begin
          if TiwUserSession(WebApplication.Data).MatrizDifVG[k, segunda] > PerdaMaximaZ[instancia] then
          begin
          PerdaMaximaZ[instancia] := MatrizDifVG[k, segunda];
          OrdemPerdaMaximaZ[instancia] := k;
          end;

          if MatrizDifVG[k, alts] > PerdaMaximaS[k] then
          begin
          PerdaMaximaS[instancia] := MatrizDifVG[k, alts];
          OrdemPerdaMaximaS[instancia] := k;
          end;



        end;

      {  if MatrizRankingAlt[k, segunda] = 0 then // Z é sol
        begin

        if MatrizDifVG[k, altx] > PerdaMaximaX[instancia] then
        begin
        PerdaMaximaX[instancia] := MatrizDifVG[k, altx];
        OrdemPerdaMaximaX[instancia] := k;
        end;

        if MatrizDifVG[k, alts] > PerdaMaximaS[instancia] then
        begin
        PerdaMaximaS[instancia] := MatrizDifVG[k, alts];
        OrdemPerdaMaximaS[instancia] := k;
        end;

        end;    }

        if MatrizRankingAlt[k, altS] = 0 then // S é sol
        begin

        if MatrizDifVG[k, altx] > PerdaMaximaX[instancia] then
        begin
        PerdaMaximaX[instancia] := MatrizDifVG[k, altx];
        OrdemPerdaMaximaX[instancia] := k;
        end;

        if MatrizDifVG[k, segunda] > PerdaMaximaZ[instancia] then
        begin
        TiwUserSession(WebApplication.Data). PerdaMaximaZ[instancia] := TiwUserSession(WebApplication.Data).MatrizDifVG[k, segunda];
        OrdemPerdaMaximaZ[instancia] := k;
        end;
        end;

      end; // for k


      // maxima diferença de consequencia normalizada para a alt X
      aux2 := 0;
      for i := 0 to numcrit - 1 do
      for j := 0 to numcrit - 1 do
      begin
        if i = j then continue;

        if (matrizconseqnorm[altx, i] - matrizconseqnorm[altx, j]) > aux2  then
        begin
          aux2 := (matrizconseqnorm[altx, i] - matrizconseqnorm[altx, j]);
          ConsequenceRangex[instancia] := aux2;
        end;
      end;
      // maxima diferença de consequencia normalizada para a alt S
      aux2 := 0;
      for i := 0 to numcrit - 1 do
      for j := 0 to numcrit - 1 do
      begin
        if i = j then continue;

        if (matrizconseqnorm[alts, i] - matrizconseqnorm[alts, j]) > aux2  then
        begin
          aux2 := (matrizconseqnorm[alts, i] - matrizconseqnorm[alts, j]);
          ConsequenceRanges[instancia] := aux2;
        end;
      end;
      // maxima diferença de consequencia normalizada para a alt Z
      aux2 := 0;
      for i := 0 to numcrit - 1 do
      for j := 0 to numcrit - 1 do
      begin
        if i = j then continue;

        if (matrizconseqnorm[segunda, i] - matrizconseqnorm[segunda, j]) > aux2  then
        begin
          aux2 := (matrizconseqnorm[segunda, i] - matrizconseqnorm[segunda, j]);
          ConsequenceRangeZ[instancia] := aux2;
        end;
      end;

      // maxima diferença de vg para a alt X
      aux3 := 0;
      aux2 := 0;
      for i := 0 to length(casesordemcrit) - 1 do
      begin
        if resultadoroc[i, altx] < aux2 then
        aux2 := resultadoroc[i, altx];

        if resultadoroc[i, altx] > aux3 then
        aux3 := resultadoroc[i, altx];
      end;
      difvgx[instancia] := aux3 - aux2;

      // maxima diferença de vg para a alt Z
      aux3 := 0;
      aux2 := 0;
      for i := 0 to length(casesordemcrit) - 1 do
      begin
        if resultadoroc[i, segunda] < aux2 then
        aux2 := resultadoroc[i, segunda];

        if resultadoroc[i, segunda] > aux3 then
        aux3 := resultadoroc[i, segunda];
      end;
      difvgz[instancia] := aux3 - aux2;

      // maxima diferença de vg para a alt s
      aux3 := 0;
      aux2 := 0;
      for i := 0 to length(casesordemcrit) - 1 do
      begin
        if resultadoroc[i, alts] < aux2 then
        aux2 := resultadoroc[i, alts];

        if resultadoroc[i, alts] > aux3 then
        aux3 := resultadoroc[i, alts];
      end;
      difvgS[instancia] := aux3 - aux2;




   for j := 0 to Numalt - 1 do
   begin
   auxprimeira[j] := 0;
   auxsegunda[j] := 0;
   auxterceira[j] := 0;
   end;
    for k := 0 to Length(CasesOrdemcrit) - 1 do
    begin
      auxprimeira[Tiwusersession(webapplication.data).MatrizRankingAlt[k, primeira]] := 1 + auxprimeira
        [MatrizRankingAlt[k, primeira]];


      auxsegunda[MatrizRankingAlt[k, segunda]] := 1 + auxsegunda
        [MatrizRankingAlt[k, segunda]];

     rankingprimeira[MatrizRankingAlt[k, primeira]] := 1 + rankingprimeira
        [MatrizRankingAlt[k, primeira]];

     rankingsegunda[MatrizRankingAlt[k, segunda]] := 1 + rankingsegunda
        [MatrizRankingAlt[k, segunda]];
    end;
    VetorAlternativaZ[instancia, 3] := (auxSegunda[0] + auxSegunda[1]) / length(casesordemcrit);

    VetorAlternativaZ[instancia, 4] := (auxSegunda[0] + auxSegunda[1] + auxSegunda[2]) / length(casesordemcrit);
    for j := 0 to 2 do
    begin

       //VetorAlternativaX[instancia, j] := auxprimeira[j] / length(casesordemcrit);

       if VetorAlternativaX[instancia, j] < 1 then
       VetorAlternativaZ[instancia, j] := auxSegunda[j] / length(casesordemcrit);

      //RankingPrimeira[j] := auxprimeira[j] + RankingPrimeira[j];
      //RankingSegunda[j] := auxsegunda[j] + RankingSegunda[j];
      //RankingTerceira[j] := auxterceira[j] + RankingTerceira[j];



      // Procedimento para estimar a antiga regra de decisão do Surrogate Input. Atualmente (17/03/2026) não está sendo utilizado
      if (auxprimeira[j] > 0.4 * Length(CasesOrdemcrit)) and (auxprimeira[j] <= 0.5 * Length(CasesOrdemcrit)) then
      begin
        MatrizFaixaPrimeira[j, 0] := MatrizFaixaPrimeira[j, 0] + 1;
      end;

      if (auxprimeira[j] > 0.5 * Length(CasesOrdemcrit)) and (auxprimeira[j] <= 0.6 * Length(CasesOrdemcrit)) then
      begin
        MatrizFaixaPrimeira[j, 1] := MatrizFaixaPrimeira[j, 1] + 1;
      end;

      if (auxprimeira[j] > 0.6 * Length(CasesOrdemcrit)) and (auxprimeira[j] <= 0.7 * Length(CasesOrdemcrit)) then
      begin
        MatrizFaixaPrimeira[j, 2] := MatrizFaixaPrimeira[j, 2] + 1;
      end;

      if (auxprimeira[j] > 0.7 * Length(CasesOrdemcrit)) and (auxprimeira[j] <= 0.8 * Length(CasesOrdemcrit)) then
      begin
        MatrizFaixaPrimeira[j, 3] := MatrizFaixaPrimeira[j, 3] + 1;
      end;

      if (auxprimeira[j] > 0.8 * Length(CasesOrdemcrit)) and (auxprimeira[j] <= 0.9 * Length(CasesOrdemcrit)) then
      begin
        MatrizFaixaPrimeira[j, 4] := MatrizFaixaPrimeira[j, 4] + 1;
      end;

      if auxprimeira[j] > 0.9 * Length(CasesOrdemcrit) then
      begin
        MatrizFaixaPrimeira[j, 5] := MatrizFaixaPrimeira[j, 5] + 1;
      end;

     { if (probalt[primeira]) < Length(CasesOrdemcrit) then
      begin
        if auxsegunda[j] > 0.4 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 0] := MatrizFaixasegunda[j, 0] + 1;
        end;

        if auxsegunda[j] > 0.5 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 1] := MatrizFaixasegunda[j, 1] + 1;
        end;

        if auxsegunda[j] > 0.6 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 2] := MatrizFaixasegunda[j, 2] + 1;
        end;

        if auxsegunda[j] > 0.7 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 3] := MatrizFaixasegunda[j, 3] + 1;
        end;

        if auxsegunda[j] > 0.8 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 4] := MatrizFaixasegunda[j, 4] + 1;
        end;

        if auxsegunda[j] > 0.9 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixasegunda[j, 5] := MatrizFaixasegunda[j, 5] + 1;
        end;
      end;

      if (probalt[primeira] + probalt[segunda]) < Length(CasesOrdemcrit) then
      begin
        if auxterceira[j] > 0.4 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 0] := MatrizFaixaterceira[j, 0] + 1;
        end;

        if auxterceira[j] > 0.5 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 1] := MatrizFaixaterceira[j, 1] + 1;
        end;

        if auxterceira[j] > 0.6 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 2] := MatrizFaixaterceira[j, 2] + 1;
        end;

        if auxterceira[j] > 0.7 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 3] := MatrizFaixaterceira[j, 3] + 1;
        end;

        if auxterceira[j] > 0.8 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 4] := MatrizFaixaterceira[j, 4] + 1;
        end;

        if auxterceira[j] > 0.9 * Length(CasesOrdemcrit) then
        begin
          MatrizFaixaterceira[j, 5] := MatrizFaixaterceira[j, 5] + 1;
        end;
      end; }

    end;


    // REGRAS DE DECISÃO

    pular1 := false;
    // REGRA 1
    if (probalt[primeira] >= 0.8 * (Length(CasesOrdemcrit))) then
    begin
      inc(regra1);
      pular1 := true;
    {
      mediaperdamedia[0, 0] := mediaperdamedia[0, 0] + perdamedia[primeira];
      mediaperdamaxima[0, 0] := mediaperdamaxima[0, 0] + perdamaxima[primeira];
      if maximaperdamedia[0, 0] < perdamedia[primeira] then
        maximaperdamedia[0, 0] := perdamedia[primeira];
      if maximaperdamaxima[0, 0] < perdamaxima[primeira] then
        maximaperdamaxima[0, 0] := perdamaxima[primeira];

      mediaperdamedia[0, 4] := mediaperdamedia[0, 4] + perdamedia[primeira];
      mediaperdamaxima[0, 4] := mediaperdamaxima[0, 4] + perdamaxima[primeira];
      mediaperdamedia[3, 0] := mediaperdamedia[3, 0] + perdamedia[primeira];
      mediaperdamaxima[3, 0] := mediaperdamaxima[3, 0] + perdamaxima[primeira];
      mediaperdamedia[3, 4] := mediaperdamedia[3, 4] + perdamedia[primeira];
      mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] + perdamaxima[primeira];

      inc(probregra[0, 0]);
      inc(probregra[3, 0]);
      inc(probregra[0, 4]);
      inc(probregra[3, 4]); }

    end;
    if (probalt[primeira] >= 0.7 * (Length(CasesOrdemcrit))) and
      (probalt[primeira] < 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
      then
    begin
      if (perdamedia[primeira] < 0.2) and (perdamaxima[primeira] < 0.4) then
      begin
        inc(regra1);
        pular1 := true;

        {mediaperdamedia[0, 1] := mediaperdamedia[0, 1] + perdamedia[primeira];
        mediaperdamaxima[0, 1] := mediaperdamaxima[0, 1] + perdamaxima
          [primeira];
        if maximaperdamedia[0, 1] < perdamedia[primeira] then
          maximaperdamedia[0, 1] := perdamedia[primeira];
        if maximaperdamaxima[0, 1] < perdamaxima[primeira] then
          maximaperdamaxima[0, 1] := perdamaxima[primeira];

        mediaperdamedia[0, 4] := mediaperdamedia[0, 4] + perdamedia[primeira];
        mediaperdamaxima[0, 4] := mediaperdamaxima[0, 4] + perdamaxima
          [primeira];
        mediaperdamedia[3, 1] := mediaperdamedia[3, 1] + perdamedia[primeira];
        mediaperdamaxima[3, 1] := mediaperdamaxima[3, 1] + perdamaxima
          [primeira];
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] + perdamedia[primeira];
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] + perdamaxima
          [primeira];

        inc(probregra[0, 1]);
        inc(probregra[3, 1]);
        inc(probregra[0, 4]);
        inc(probregra[3, 4]); }

      end;
    end;
    if (probalt[primeira] >= 0.6 * (Length(CasesOrdemcrit))) and
      (probalt[primeira] < 0.7 * (Length(CasesOrdemcrit))) and (pular1 = false)
      then
    begin
      if (perdamedia[primeira] < 0.15) and (perdamaxima[primeira] < 0.30) then
      begin
        inc(regra1);
        pular1 := true;
       {
        mediaperdamedia[0, 2] := mediaperdamedia[0, 2] + perdamedia[primeira];
        mediaperdamaxima[0, 2] := mediaperdamaxima[0, 2] + perdamaxima
          [primeira];
        if maximaperdamedia[0, 2] < perdamedia[primeira] then
          maximaperdamedia[0, 2] := perdamedia[primeira];
        if maximaperdamaxima[0, 2] < perdamaxima[primeira] then
          maximaperdamaxima[0, 2] := perdamaxima[primeira];

        mediaperdamedia[0, 4] := mediaperdamedia[0, 4] + perdamedia[primeira];
        mediaperdamaxima[0, 4] := mediaperdamaxima[0, 4] + perdamaxima
          [primeira];
        mediaperdamedia[3, 2] := mediaperdamedia[3, 2] + perdamedia[primeira];
        mediaperdamaxima[3, 2] := mediaperdamaxima[3, 2] + perdamaxima
          [primeira];
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] + perdamedia[primeira];
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] + perdamaxima
          [primeira];

        inc(probregra[0, 2]);
        inc(probregra[3, 2]);
        inc(probregra[0, 4]);
        inc(probregra[3, 4]);    }

      end;
    end;
    if (probalt[primeira] >= 0.5 * (Length(CasesOrdemcrit))) and
      (probalt[primeira] < 0.6 * (Length(CasesOrdemcrit))) and (pular1 = false)
      then
    begin
      if (perdamedia[primeira] < 0.10) and (perdamaxima[primeira] < 0.20) then
      begin
        inc(regra1);
        pular1 := true;

       { mediaperdamedia[0, 3] := mediaperdamedia[0, 3] + perdamedia[primeira];
        mediaperdamaxima[0, 3] := mediaperdamaxima[0, 3] + perdamaxima
          [primeira];
        if maximaperdamedia[0, 3] < perdamedia[primeira] then
          maximaperdamedia[0, 3] := perdamedia[primeira];
        if maximaperdamaxima[0, 3] < perdamaxima[primeira] then
          maximaperdamaxima[0, 3] := perdamaxima[primeira];

        mediaperdamedia[0, 4] := mediaperdamedia[0, 4] + perdamedia[primeira];
        mediaperdamaxima[0, 4] := mediaperdamaxima[0, 4] + perdamaxima
          [primeira];
        mediaperdamedia[3, 3] := mediaperdamedia[3, 3] + perdamedia[primeira];
        mediaperdamaxima[3, 3] := mediaperdamaxima[3, 3] + perdamaxima
          [primeira];
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] + perdamedia[primeira];
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] + perdamaxima
          [primeira];

        inc(probregra[0, 3]);
        inc(probregra[3, 3]);
        inc(probregra[0, 4]);
        inc(probregra[3, 4]);          }
      end;
    end;

    // REGRA 2
    auxresultado := TiwUserSession(WebApplication.Data).probalt[primeira]
      + probalt[segunda];
    if (auxresultado >= 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
      then
    begin
      inc(regra2);
      pular1 := true;
      {mediaperdamedia[1, 0] := mediaperdamedia[1, 0] +
        (perdamedia[primeira] + perdamedia[segunda]) / 2;
      mediaperdamaxima[1, 0] := mediaperdamaxima[1, 0] +
        (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
      if maximaperdamedia[1, 0] < perdamedia[primeira] then
        maximaperdamedia[1, 0] := perdamedia[primeira];
      if maximaperdamaxima[1, 0] < perdamaxima[primeira] then
        maximaperdamaxima[1, 0] := perdamaxima[primeira];
      if maximaperdamedia[1, 0] < perdamedia[segunda] then
        maximaperdamedia[1, 0] := perdamedia[segunda];
      if maximaperdamaxima[1, 0] < perdamaxima[segunda] then
        maximaperdamaxima[1, 0] := perdamaxima[segunda];

      mediaperdamedia[1, 4] := mediaperdamedia[1, 4] +
        (perdamedia[primeira] + perdamedia[segunda]) / 2;
      mediaperdamaxima[1, 4] := mediaperdamaxima[1, 4] +
        (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
      mediaperdamedia[3, 0] := mediaperdamedia[3, 0] +
        (perdamedia[primeira] + perdamedia[segunda]) / 2;
      mediaperdamaxima[3, 0] := mediaperdamaxima[3, 0] +
        (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
      mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
        (perdamedia[primeira] + perdamedia[segunda]) / 2;
      mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
        (perdamaxima[primeira] + perdamaxima[segunda]) / 2;

      inc(probregra[1, 0]);
      inc(probregra[3, 0]);
      inc(probregra[1, 4]);
      inc(probregra[3, 4]); }

    end;
    if (auxresultado >= 0.7 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.2) and (perdamaxima[primeira] < 0.4) and
        (perdamedia[segunda] < 0.2) and (perdamaxima[segunda] < 0.4) then
      begin
        inc(regra2);
        pular1 := true;

       { mediaperdamedia[1, 1] := mediaperdamedia[1, 1] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 1] := mediaperdamaxima[1, 1] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        if maximaperdamedia[1, 1] < perdamedia[primeira] then
          maximaperdamedia[1, 1] := perdamedia[primeira];
        if maximaperdamaxima[1, 1] < perdamaxima[primeira] then
          maximaperdamaxima[1, 1] := perdamaxima[primeira];
        if maximaperdamedia[1, 1] < perdamedia[segunda] then
          maximaperdamedia[1, 1] := perdamedia[segunda];
        if maximaperdamaxima[1, 1] < perdamaxima[segunda] then
          maximaperdamaxima[1, 1] := perdamaxima[segunda];

        mediaperdamedia[1, 4] := mediaperdamedia[1, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 4] := mediaperdamaxima[1, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 1] := mediaperdamedia[3, 1] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 1] := mediaperdamaxima[3, 1] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;

        inc(probregra[1, 1]);
        inc(probregra[3, 1]);
        inc(probregra[1, 4]);
        inc(probregra[3, 4]);     }
      end;
    end;
    if (auxresultado >= 0.6 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.7 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.15) and (perdamaxima[primeira] < 0.30) and
        (perdamedia[segunda] < 0.15) and (perdamaxima[segunda] < 0.3) then
      begin
        inc(regra2);
        pular1 := true;

       { mediaperdamedia[1, 2] := mediaperdamedia[1, 2] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 2] := mediaperdamaxima[1, 2] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        if maximaperdamedia[1, 2] < perdamedia[primeira] then
          maximaperdamedia[1, 2] := perdamedia[primeira];
        if maximaperdamaxima[1, 2] < perdamaxima[primeira] then
          maximaperdamaxima[1, 2] := perdamaxima[primeira];
        if maximaperdamedia[1, 2] < perdamedia[segunda] then
          maximaperdamedia[1, 2] := perdamedia[segunda];
        if maximaperdamaxima[1, 2] < perdamaxima[segunda] then
          maximaperdamaxima[1, 2] := perdamaxima[segunda];

        mediaperdamedia[1, 4] := mediaperdamedia[1, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 4] := mediaperdamaxima[1, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 2] := mediaperdamedia[3, 2] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 2] := mediaperdamaxima[3, 2] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;

        inc(probregra[1, 2]);
        inc(probregra[3, 2]);
        inc(probregra[1, 4]);
        inc(probregra[3, 4]);                    }
      end;
    end;
    if (auxresultado >= 0.5 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.6 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.10) and (perdamaxima[primeira] < 0.20) and
        (perdamedia[segunda] < 0.1) and (perdamaxima[segunda] < 0.2) then
      begin
        inc(regra2);
        pular1 := true;
     {
        mediaperdamedia[1, 3] := mediaperdamedia[1, 3] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 3] := mediaperdamaxima[1, 3] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        if maximaperdamedia[1, 3] < perdamedia[primeira] then
          maximaperdamedia[1, 3] := perdamedia[primeira];
        if maximaperdamaxima[1, 3] < perdamaxima[primeira] then
          maximaperdamaxima[1, 3] := perdamaxima[primeira];
        if maximaperdamedia[1, 3] < perdamedia[segunda] then
          maximaperdamedia[1, 3] := perdamedia[segunda];
        if maximaperdamaxima[1, 3] < perdamaxima[segunda] then
          maximaperdamaxima[1, 3] := perdamaxima[segunda];

        mediaperdamedia[1, 4] := mediaperdamedia[1, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[1, 4] := mediaperdamaxima[1, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 3] := mediaperdamedia[3, 3] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 3] := mediaperdamaxima[3, 3] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda]) / 2;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda]) / 2;

        inc(probregra[1, 3]);
        inc(probregra[3, 3]);
        inc(probregra[1, 4]);
        inc(probregra[3, 4]);     }
      end;
    end;

    // REGRA 3
    auxresultado := probalt[primeira] + probalt[segunda] + probalt[terceira];
    if (auxresultado >= 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
      then
    begin
      inc(regra3);
      pular1 := true;

     { mediaperdamedia[2, 0] := mediaperdamedia[2, 0] +
        (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira]) / 3;
      mediaperdamaxima[2, 0] := mediaperdamaxima[2, 0] +
        (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
        / 3;
      if maximaperdamedia[2, 0] < perdamedia[primeira] then
        maximaperdamedia[2, 0] := perdamedia[primeira];
      if maximaperdamaxima[2, 0] < perdamaxima[primeira] then
        maximaperdamaxima[2, 0] := perdamaxima[primeira];
      if maximaperdamedia[2, 0] < perdamedia[segunda] then
        maximaperdamedia[2, 0] := perdamedia[segunda];
      if maximaperdamaxima[2, 0] < perdamaxima[segunda] then
        maximaperdamaxima[2, 0] := perdamaxima[segunda];
      if maximaperdamedia[2, 0] < perdamedia[terceira] then
        maximaperdamedia[2, 0] := perdamedia[terceira];
      if maximaperdamaxima[2, 0] < perdamaxima[terceira] then
        maximaperdamaxima[2, 0] := perdamaxima[terceira];

      mediaperdamedia[2, 4] := mediaperdamedia[2, 4] +
        (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira]) / 3;
      mediaperdamaxima[2, 4] := mediaperdamaxima[2, 4] +
        (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
        / 3;
      mediaperdamedia[3, 0] := mediaperdamedia[3, 0] +
        (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira]) / 3;
      mediaperdamaxima[3, 0] := mediaperdamaxima[3, 0] +
        (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
        / 3;
      mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
        (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira]) / 3;
      mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
        (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
        / 3;

      inc(probregra[2, 0]);
      inc(probregra[3, 0]);
      inc(probregra[2, 4]);
      inc(probregra[3, 4]);   }

    end;
    if (auxresultado >= 0.7 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.2) and (perdamaxima[primeira] < 0.4) and
        (perdamedia[segunda] < 0.2) and (perdamaxima[segunda] < 0.4) and
        (perdamedia[terceira] < 0.2) and (perdamaxima[terceira] < 0.4) then
      begin
        inc(regra3);
        pular1 := true;

       { mediaperdamedia[2, 1] := mediaperdamedia[2, 1] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 1] := mediaperdamaxima[2, 1] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        if maximaperdamedia[2, 1] < perdamedia[primeira] then
          maximaperdamedia[2, 1] := perdamedia[primeira];
        if maximaperdamaxima[2, 1] < perdamaxima[primeira] then
          maximaperdamaxima[2, 1] := perdamaxima[primeira];
        if maximaperdamedia[2, 1] < perdamedia[segunda] then
          maximaperdamedia[2, 1] := perdamedia[segunda];
        if maximaperdamaxima[2, 1] < perdamaxima[segunda] then
          maximaperdamaxima[2, 1] := perdamaxima[segunda];
        if maximaperdamedia[2, 1] < perdamedia[terceira] then
          maximaperdamedia[2, 1] := perdamedia[terceira];
        if maximaperdamaxima[2, 1] < perdamaxima[terceira] then
          maximaperdamaxima[2, 1] := perdamaxima[terceira];

        mediaperdamedia[2, 4] := mediaperdamedia[2, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 4] := mediaperdamaxima[2, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 1] := mediaperdamedia[3, 1] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 1] := mediaperdamaxima[3, 1] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;

        inc(probregra[2, 1]);
        inc(probregra[3, 1]);
        inc(probregra[2, 4]);
        inc(probregra[3, 4]);   }
      end;
    end;
    if (auxresultado >= 0.6 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.7 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.15) and (perdamaxima[primeira] < 0.30) and
        (perdamedia[segunda] < 0.15) and (perdamaxima[segunda] < 0.3) and
        (perdamedia[terceira] < 0.15) and (perdamaxima[terceira] < 0.3) then
      begin
        inc(regra3);
        pular1 := true;

       { mediaperdamedia[2, 2] := mediaperdamedia[2, 2] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 2] := mediaperdamaxima[2, 2] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        if maximaperdamedia[2, 2] < perdamedia[primeira] then
          maximaperdamedia[2, 2] := perdamedia[primeira];
        if maximaperdamaxima[2, 2] < perdamaxima[primeira] then
          maximaperdamaxima[2, 2] := perdamaxima[primeira];
        if maximaperdamedia[2, 2] < perdamedia[segunda] then
          maximaperdamedia[2, 2] := perdamedia[segunda];
        if maximaperdamaxima[2, 2] < perdamaxima[segunda] then
          maximaperdamaxima[2, 2] := perdamaxima[segunda];
        if maximaperdamedia[2, 2] < perdamedia[terceira] then
          maximaperdamedia[2, 2] := perdamedia[terceira];
        if maximaperdamaxima[2, 2] < perdamaxima[terceira] then
          maximaperdamaxima[2, 2] := perdamaxima[terceira];

        mediaperdamedia[2, 4] := mediaperdamedia[2, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 4] := mediaperdamaxima[2, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 2] := mediaperdamedia[3, 2] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 2] := mediaperdamaxima[3, 2] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;

        inc(probregra[2, 2]);
        inc(probregra[3, 2]);
        inc(probregra[2, 4]);
        inc(probregra[3, 4]);  }
      end;
    end;
    if (auxresultado >= 0.5 * (Length(CasesOrdemcrit))) and
      (auxresultado < 0.6 * (Length(CasesOrdemcrit))) and (pular1 = false) then
    begin
      if (perdamedia[primeira] < 0.10) and (perdamaxima[primeira] < 0.20) and
        (perdamedia[segunda] < 0.1) and (perdamaxima[segunda] < 0.2) and
        (perdamedia[terceira] < 0.1) and (perdamaxima[terceira] < 0.2) then
      begin
        inc(regra3);
        pular1 := true;

        {mediaperdamedia[2, 3] := mediaperdamedia[2, 3] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 3] := mediaperdamaxima[2, 3] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        if maximaperdamedia[2, 3] < perdamedia[primeira] then
          maximaperdamedia[2, 3] := perdamedia[primeira];
        if maximaperdamaxima[2, 3] < perdamaxima[primeira] then
          maximaperdamaxima[2, 3] := perdamaxima[primeira];
        if maximaperdamedia[2, 3] < perdamedia[segunda] then
          maximaperdamedia[2, 3] := perdamedia[segunda];
        if maximaperdamaxima[2, 3] < perdamaxima[segunda] then
          maximaperdamaxima[2, 3] := perdamaxima[segunda];
        if maximaperdamedia[2, 3] < perdamedia[terceira] then
          maximaperdamedia[2, 3] := perdamedia[terceira];
        if maximaperdamaxima[2, 3] < perdamaxima[terceira] then
          maximaperdamaxima[2, 3] := perdamaxima[terceira];

        mediaperdamedia[2, 4] := mediaperdamedia[2, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[2, 4] := mediaperdamaxima[2, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 3] := mediaperdamedia[3, 3] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 3] := mediaperdamaxima[3, 3] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamedia[3, 4] := mediaperdamedia[3, 4] +
          (perdamedia[primeira] + perdamedia[segunda] + perdamedia[terceira])
          / 3;
        mediaperdamaxima[3, 4] := mediaperdamaxima[3, 4] +
          (perdamaxima[primeira] + perdamaxima[segunda] + perdamedia[terceira])
          / 3;

        inc(probregra[2, 3]);
        inc(probregra[3, 3]);
        inc(probregra[2, 4]);
        inc(probregra[3, 4]);  }
      end;
    end;

    // REGRA 4
    if pular1 = false then
      inc(regra4);

    if sobreclassificacao then
    begin
      // Promethee - Caso do Surrogate

      ComparacaoParAPar;
      auxpromethee := k;
      CalculoFluxos;

      for j := 0 to Numalt - 1 do
      begin
        TiwUserSession(WebApplication.Data).ResultadoPromethee
          [Length(CasesOrdemcrit) - 1, j] := NetFlow[j];
      end;
      k := Length(CasesOrdemcrit) - 1;
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if ResultadoPromethee[k, j] > auxresultado then
          auxresultado := ResultadoPromethee[k, j];
      end;
      for j := 0 to Numalt - 1 do
      begin
        TiwUserSession(WebApplication.Data).MatrizDifVGPromethee[k, j] :=
          auxresultado - ResultadoPromethee[k, j];
        if auxresultado = ResultadoPromethee[k, j] then
        begin
          TiwUserSession(WebApplication.Data).MatrizPOAPromethee[k, j] := 1;
          TiwUserSession(WebApplication.Data).probaltPromethee[j] :=
            probaltPromethee[j] + 1;
        end;
      end;
      for j := 0 to Numalt - 1 do
      begin
        if perdamaximaPromethee[j] < MatrizDifVGPromethee[k, j] then
          perdamaximaPromethee[j] := MatrizDifVGPromethee[k, j];
        perdamediaPromethee[j] := perdamediaPromethee[j] + MatrizDifVGPromethee
          [k, j];
      end;
      // Fim do Promethee - Caso do Surrogate
      Promethee; // Armazena as soluções para o Promethee

      for j := 0 to Numalt - 1 do
      begin
        if probaltPromethee[j] < Length(CasesOrdemcrit) then
          perdamediaPromethee[j] := perdamediaPromethee[j] /
            (Length(CasesOrdemcrit) - probaltPromethee[j]);
        if probaltPromethee[j] = Length(CasesOrdemcrit) then
          perdamediaPromethee[j] := 0;
      end;

      auxresultado := 0;
      primeira := 0;
      for j := 0 to Numalt - 1 do
      begin
        if probaltPromethee[j] > auxresultado then
        begin
          auxresultado := probaltPromethee[j];
          primeira := j;
        end;
      end;
      auxresultado := 0;
      segunda := 0;
      for j := 0 to Numalt - 1 do
      begin
        if (probaltPromethee[j] > auxresultado) and (primeira <> j) then
        begin
          auxresultado := probaltPromethee[j];
          segunda := j;
        end;
      end;
      auxresultado := 0;
      terceira := 0;
      for j := 0 to Numalt - 1 do
      begin
        if (probaltPromethee[j] > auxresultado) and (primeira <> j) and
          (segunda <> j) then
        begin
          auxresultado := probaltPromethee[j];
          terceira := j;
        end;
      end;

      // REGRAS DE DECISÃO
      pular1 := false;
      // REGRA 1
      if (probaltPromethee[primeira] >= 0.8 * (Length(CasesOrdemcrit))) then
      begin
        inc(regra1);
        pular1 := true;
      end;
      if (probaltPromethee[primeira] >= 0.7 * (Length(CasesOrdemcrit))) and
        (probaltPromethee[primeira] < 0.8 * (Length(CasesOrdemcrit))) and
        (pular1 = false) then
      begin
        if (perdamediaPromethee[primeira] < 0.2) and
          (perdamaximaPromethee[primeira] < 0.4) then
        begin
          inc(regra1Promethee);
          pular1 := true;
        end;
      end;
      if (probaltPromethee[primeira] >= 0.6 * (Length(CasesOrdemcrit))) and
        (probaltPromethee[primeira] < 0.7 * (Length(CasesOrdemcrit))) and
        (pular1 = false) then
      begin
        if (perdamediaPromethee[primeira] < 0.15) and
          (perdamaximaPromethee[primeira] < 0.30) then
        begin
          inc(regra1Promethee);
          pular1 := true;
        end;
      end;
      if (probaltPromethee[primeira] >= 0.5 * (Length(CasesOrdemcrit))) and
        (probaltPromethee[primeira] < 0.6 * (Length(CasesOrdemcrit))) and
        (pular1 = false) then
      begin
        if (perdamediaPromethee[primeira] < 0.10) and
          (perdamaximaPromethee[primeira] < 0.20) then
        begin
          inc(regra1Promethee);
          pular1 := true;
        end;
      end;

      // REGRA 2
      auxresultado := TiwUserSession(WebApplication.Data).probaltPromethee
        [primeira] + probaltPromethee[segunda];
      if (auxresultado >= 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        inc(regra2Promethee);
        pular1 := true;
      end;
      if (auxresultado >= 0.7 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.2) and
          (perdamaximaPromethee[primeira] < 0.4) and
          (perdamediaPromethee[segunda] < 0.2) and
          (perdamaximaPromethee[segunda] < 0.4) then
        begin
          inc(regra2Promethee);
          pular1 := true;
        end;
      end;
      if (auxresultado >= 0.6 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.7 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.15) and
          (perdamaximaPromethee[primeira] < 0.30) and
          (perdamediaPromethee[segunda] < 0.15) and
          (perdamaximaPromethee[segunda] < 0.3) then
        begin
          inc(regra2Promethee);
          pular1 := true;
        end;
      end;
      if (auxresultado >= 0.5 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.6 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.10) and
          (perdamaximaPromethee[primeira] < 0.20) and
          (perdamediaPromethee[segunda] < 0.1) and
          (perdamaximaPromethee[segunda] < 0.2) then
        begin
          inc(regra2Promethee);
          pular1 := true;
        end;
      end;

      // REGRA 3
      auxresultado := probaltPromethee[primeira] + probaltPromethee[segunda]
        + probaltPromethee[terceira];
      if (auxresultado >= 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        inc(regra3Promethee);
        pular1 := true;
      end;
      if (auxresultado >= 0.7 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.8 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.2) and
          (perdamaximaPromethee[primeira] < 0.4) and
          (perdamediaPromethee[segunda] < 0.2) and
          (perdamaximaPromethee[segunda] < 0.4) and
          (perdamediaPromethee[terceira] < 0.2) and
          (perdamaximaPromethee[terceira] < 0.4) then
        begin
          inc(regra3Promethee);
          pular1 := true;
        end;
      end;
      if (auxresultado >= 0.6 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.7 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.15) and
          (perdamaximaPromethee[primeira] < 0.30) and
          (perdamediaPromethee[segunda] < 0.15) and
          (perdamaximaPromethee[segunda] < 0.3) and
          (perdamediaPromethee[terceira] < 0.15) and
          (perdamaximaPromethee[terceira] < 0.3) then
        begin
          inc(regra3Promethee);
          pular1 := true;
        end;
      end;
      if (auxresultado >= 0.5 * (Length(CasesOrdemcrit))) and
        (auxresultado < 0.6 * (Length(CasesOrdemcrit))) and (pular1 = false)
        then
      begin
        if (perdamediaPromethee[primeira] < 0.10) and
          (perdamaximaPromethee[primeira] < 0.20) and
          (perdamediaPromethee[segunda] < 0.1) and
          (perdamaximaPromethee[segunda] < 0.2) and
          (perdamediaPromethee[terceira] < 0.1) and
          (perdamaximaPromethee[terceira] < 0.2) then
        begin
          inc(regra3Promethee);
          pular1 := true;
        end;
      end;

      // REGRA 4
      if pular1 = false then
        inc(regra4Promethee);
    end;

    if aditivo then
    begin
      SetLength(MatrizSol, 1, Numalt);
      SetLength(ResultAlt, Numalt);
      SetLength(ResultSol, 1);
      for j := 0 to Numalt - 1 do
        MatrizSol[0, j] := MatrizPOA[0, j];
      ResultSol[0] := 1;

      CaseSol[0] := 1;

      for I := 1 to Length(MatrizPOA) - 1 do
      begin

        k := 0;
        NovaSol := true;
        while k <= Length(MatrizSol) - 1 do
        begin
          aux := 0;
          for j := 0 to Numalt - 1 do
          begin
            if MatrizSol[k, j] = MatrizPOA[I, j] then
              inc(aux);
          end;
          if aux = Numalt then
          begin
            ResultSol[k] := ResultSol[k] + 1;
            CaseSol[I] := k + 1;
            k := Length(MatrizSol);
            NovaSol := false;
          end
          else
          begin
            inc(k);
          end;
        end; // k
        if NovaSol = true then
        begin
          SetLength(TiwUserSession(WebApplication.Data).MatrizSol, Length
              (MatrizSol) + 1, Numalt);
          SetLength(TiwUserSession(WebApplication.Data).ResultSol, Length
              (MatrizSol));
          for j := 0 to Numalt - 1 do
            MatrizSol[Length(MatrizSol) - 1, j] := MatrizPOA[I, j];
          ResultSol[Length(MatrizSol) - 1] := 1;
          CaseSol[I] := Length(MatrizSol);
        end;

      end; // i


 {     for I := 0 to Numalt - 1 do
      begin
        TiwUserSession(WebApplication.Data).perdamediaindividuos[I] := 0;
        TiwUserSession(WebApplication.Data).perdamaximaindividuos[I] := 0;
        if TiwUserSession(WebApplication.Data).probalt[I] > 0 then
        begin
          for k := 0 to Length(MatrizPOA) - 1 do
          begin
            if MatrizDifVG[k, I] <> 0 then
              perdamediaindividuos[I] := perdamediaindividuos[I]
                + TiwUserSession(WebApplication.Data).MatrizDifVG[k, I];

            if MatrizDifVG[k, I] = 0 then
            begin
              for j := 0 to Numalt - 1 do
              begin
                if I = j then
                  continue;
                ganhomedioindividuos[I] := ganhomedioindividuos[I] + MatrizDifVG
                  [k, j]

              end;
            end;
            if perdamaximaindividuos[I] < MatrizDifVG[k, I] then
              perdamaximaindividuos[I] := MatrizDifVG[k, I];

          end;
          TiwUserSession(WebApplication.Data).perdamediaindividuos[I] :=
            perdamediaindividuos[I] / Length(MatrizPOA);
          TiwUserSession(WebApplication.Data).ganhomedioindividuos[I] :=
            ganhomedioindividuos[I] / ((Numalt - 1) * Length(MatrizPOA));

          desempenhoindividuos[I] :=
            (ganhomedioindividuos[I] * ((probalt[I] / (Length(CasesOrdemcrit))))
            ) - (perdamediaindividuos[I] *
              (1 - (probalt[I] / (Length(CasesOrdemcrit)))));
        end;
      end;
    end;
    // supondo que matrizprobalt já foi declarado como: array of Double
    if Length(desempenhoindividuos) <= 1 then
      Exit;

    for I := 0 to Length(desempenhoindividuos) - 2 do
    begin
      minIdx := I;
      for j := I + 1 to Length(desempenhoindividuos) - 1 do
        if desempenhoindividuos[j] > desempenhoindividuos[minIdx] then
          minIdx := j;

      if minIdx <> I then
      begin
        tmp := desempenhoindividuos[I];
        desempenhoindividuos[I] := desempenhoindividuos[minIdx];
        TiwUserSession(WebApplication.Data).desempenhoindividuos[minIdx] := tmp;
      end;
    end;

    TiwUserSession(WebApplication.Data).desempenhopopulacao[instancia] :=
      desempenhoindividuos[0];
    mediadesempenho := mediadesempenho + desempenhopopulacao[instancia];

    if maximodesempenho < desempenhopopulacao[instancia] then
      maximodesempenho := desempenhopopulacao[instancia];

    TiwUserSession(WebApplication.Data).percentualdesempenho
      [round(100 * desempenhoindividuos[0] + 100)] := percentualdesempenho
      [round(100 * desempenhoindividuos[0] + 100)] + 1;

   }
    end;



  end; // with
end;

procedure TIWForm1.Promethee;
var
  I, j, l, c, xcrit, k, aux, t: Integer;
  soma, Min, auxresultado: real;
  NovaSol: Boolean;
  pular: array of Boolean;
  pular2: Boolean;
  Fatorial: real;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    finalize(MatrizSolPromethee);
    finalize(ResultAltPromethee);
    finalize(ResultSolPromethee);
    SetLength(TiwUserSession(WebApplication.Data).MatrizSolPromethee, 1,
      Numalt);
    SetLength(TiwUserSession(WebApplication.Data).ResultAltPromethee, Numalt);
    SetLength(TiwUserSession(WebApplication.Data).ResultSolPromethee, 1);
    for j := 0 to Numalt - 1 do
      MatrizSolPromethee[0, j] := MatrizPOAPromethee[0, j];
    ResultSolPromethee[0] := 1;

    CaseSolPromethee[0] := 1;

    for I := 1 to Length(MatrizPOAPromethee) - 1 do
    begin

      k := 0;
      NovaSol := true;
      while k <= Length(MatrizSolPromethee) - 1 do
      begin
        aux := 0;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSolPromethee[k, j] = MatrizPOAPromethee[I, j] then
            inc(aux);
        end;
        if aux = Numalt then
        begin
          ResultSolPromethee[k] := ResultSolPromethee[k] + 1;
          CaseSolPromethee[I] := k + 1;
          k := Length(MatrizSolPromethee);
          NovaSol := false;
        end
        else
        begin
          inc(k);
        end;
      end; // k
      if NovaSol = true then
      begin
        SetLength(TiwUserSession(WebApplication.Data).MatrizSolPromethee, Length
            (MatrizSolPromethee) + 1, Numalt);
        SetLength(TiwUserSession(WebApplication.Data).ResultSolPromethee, Length
            (MatrizSolPromethee));
        for j := 0 to Numalt - 1 do
          MatrizSolPromethee[Length(MatrizSolPromethee) - 1, j] :=
            MatrizPOAPromethee[I, j];
        ResultSolPromethee[Length(MatrizSolPromethee) - 1] := 1;
        TiwUserSession(WebApplication.Data).CaseSolPromethee[I] := Length
          (MatrizSolPromethee);
      end;

    end; // i
    SetLength(TiwUserSession(WebApplication.Data).ResultAltPromethee, Numalt);
    for j := 0 to Numalt - 1 do
    begin
      ResultAltPromethee[j] := 0;
      for I := 0 to Length(MatrizPOAPromethee) - 1 do
      begin
        ResultAltPromethee[j] := ResultAltPromethee[j] + MatrizPOAPromethee
          [I, j];
      end;
    end;

    SetLength(pular, Length(MatrizSolPromethee));

    for I := 0 to Length(MatrizSolPromethee) - 1 do
      pular[I] := false;

    for I := 0 to Length(MatrizSolPromethee) - 1 do
    begin
      auxresultado := 0;
      for j := 0 to Length(MatrizSolPromethee) - 1 do
      begin
        if pular[j] = true then
          continue;
        if TiwUserSession(WebApplication.Data).ResultSolPromethee[j]
          > auxresultado then
          auxresultado := ResultSolPromethee[j];
      end;
      pular2 := false;
      for j := 0 to Length(MatrizSolPromethee) - 1 do
      begin
        if (pular2 = true) or (pular[j] = true) then
          continue;
        if auxresultado = ResultSolPromethee[j] then
        begin
          pular[j] := true;
          pular2 := true;
        end;

      end;
    end;

  end;
end;

procedure TIWForm1.ComparacaoParAPar;
var
  xcrit, l, c: Integer;
begin
  finalize(TiwUserSession(WebApplication.Data).matrizParaPar);
  SetLength(TiwUserSession(WebApplication.Data).matrizParaPar, TiwUserSession
      (WebApplication.Data).numcrit, TiwUserSession(WebApplication.Data)
      .Numalt, TiwUserSession(WebApplication.Data).Numalt);
  xcrit := 0;
  for xcrit := 1 to TiwUserSession(WebApplication.Data).numcrit do
    for l := 1 to TiwUserSession(WebApplication.Data).Numalt - 1 do
      for c := 1 to TiwUserSession(WebApplication.Data).Numalt - 1 do
        TiwUserSession(WebApplication.Data)
          .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;

  for xcrit := 1 to TiwUserSession(WebApplication.Data).numcrit do
  begin
    if (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 1) or
      (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 3) or
      (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 5) then
    begin
      for l := 1 to TiwUserSession(WebApplication.Data).Numalt do
      begin
        for c := 1 to TiwUserSession(WebApplication.Data).Numalt do
        begin
          TiwUserSession(WebApplication.Data).comp1 := TiwUserSession
            (WebApplication.Data).MatrizConseq[c - 1, xcrit - 1];
          { Alternativa a(j) }
          TiwUserSession(WebApplication.Data).comp2 := TiwUserSession
            (WebApplication.Data).MatrizConseq[l - 1, xcrit - 1];
          { Alternativa a(i) }
          if not(c = l) then
            if (TiwUserSession(WebApplication.Data).comp2 - TiwUserSession
                (WebApplication.Data).comp1 > 0) then
              TiwUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 1
              { preenche a matriz Par a Par do critério }
            else
              TiwUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;
          { preenche a matriz Par a Par do critério }
        end;
      end;
    end;
    if (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 0) or
      (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 2) or
      (TiwUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 4) then
    begin

      for l := 1 to TiwUserSession(WebApplication.Data).Numalt do
      begin
        for c := 1 to TiwUserSession(WebApplication.Data).Numalt do
        begin
          TiwUserSession(WebApplication.Data).comp1 := TiwUserSession
            (WebApplication.Data).MatrizConseq[c - 1, xcrit - 1];
          { Alternativa a(j) }
          TiwUserSession(WebApplication.Data).comp2 := TiwUserSession
            (WebApplication.Data).MatrizConseq[l - 1, xcrit - 1];
          { Alternativa a(i) }
          if not(c = l) then
            if (TiwUserSession(WebApplication.Data).comp1 - TiwUserSession
                (WebApplication.Data).comp2 > 0) then
              TiwUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 1
              { preenche a matriz Par a Par do critério }
            else
              TiwUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;
          { preenche a matriz Par a Par do critério }
        end;
      end;
    end;
  end;
end;

procedure TIWForm1.CalculoFluxos;
var
  I, j, l, c, xcrit: Integer;
  soma, Min: real;

begin
  With TiwUserSession(WebApplication.Data) do
  begin
    // matriz grau de sobreclassificação
    finalize(TiwUserSession(WebApplication.Data).SobClassMatrix);
    finalize(TiwUserSession(WebApplication.Data).PositiveFlow);
    finalize(TiwUserSession(WebApplication.Data).NegativeFlow);
    finalize(TiwUserSession(WebApplication.Data).NetFlow);

    SetLength(TiwUserSession(WebApplication.Data).SobClassMatrix, TiwUserSession
        (WebApplication.Data).Numalt, TiwUserSession(WebApplication.Data)
        .Numalt);

    for I := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
      for j := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
        TiwUserSession(WebApplication.Data).SobClassMatrix[I, j] := 0;

    for I := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
    begin
      for j := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
      begin
        for xcrit := 0 to TiwUserSession(WebApplication.Data).numcrit - 1 do
        begin
          TiwUserSession(WebApplication.Data).SobClassMatrix[I, j] :=
            TiwUserSession(WebApplication.Data).SobClassMatrix[I, j]
            + TiwUserSession(WebApplication.Data).PesoCrit[auxpromethee, xcrit]
            * TiwUserSession(WebApplication.Data).matrizParaPar[xcrit, I, j];
        end;
      end;
    end;

    // fluxo positivo proposto por Brans e Mareschal (2002)

    SetLength(TiwUserSession(WebApplication.Data).PositiveFlow, TiwUserSession
        (WebApplication.Data).Numalt);

    for l := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
    { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
    begin
      soma := 0;
      for c := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
      begin
        soma := soma + TiwUserSession(WebApplication.Data).SobClassMatrix[l, c];
      end;
      TiwUserSession(WebApplication.Data).PositiveFlow[l] := soma /
        (TiwUserSession(WebApplication.Data).Numalt - 1);
      { FLUXO POSITIVO - preenchimento }
    end;

    // fluxo negativo

    SetLength(TiwUserSession(WebApplication.Data).NegativeFlow, TiwUserSession
        (WebApplication.Data).Numalt);

    for c := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
    { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
    begin
      soma := 0;
      for l := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
      begin
        soma := soma + TiwUserSession(WebApplication.Data).SobClassMatrix[l, c];
      end;
      TiwUserSession(WebApplication.Data).NegativeFlow[c] := soma /
        (TiwUserSession(WebApplication.Data).Numalt - 1);
    end;

    // fluxo líquido

    SetLength(TiwUserSession(WebApplication.Data).NetFlow, TiwUserSession
        (WebApplication.Data).Numalt);

    for c := 0 to TiwUserSession(WebApplication.Data).Numalt - 1 do
    { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
    begin
      TiwUserSession(WebApplication.Data).NetFlow[c] := TiwUserSession
        (WebApplication.Data).PositiveFlow[c] - TiwUserSession
        (WebApplication.Data).NegativeFlow[c];
    end;
  end;

end;

procedure TIWForm1.GerarCases;
var
  Vetor, Indices: array of Integer;
  Fatorial, Linha, I, j, k, kk, Temp: Integer;
  N: Integer;
  Matriz: array of array of Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    // Calcular o fatorial de N (número total de permutações)
    Fatorial := 1;
    N := numcrit;
    for I := 2 to N do
      Fatorial := Fatorial * I;

    // Redimensionar a matriz para armazenar todas as permutações
    SetLength(Matriz, Fatorial, N);

    // Inicializar o vetor de entrada e índices
    SetLength(Vetor, N);
    for I := 0 to N - 1 do
      Vetor[I] := I + 1;

    SetLength(Indices, N);
    for I := 0 to N - 1 do
      Indices[I] := 0;

    // Copiar a primeira permutação
    Linha := 0;
    for I := 0 to N - 1 do
      Matriz[Linha][I] := Vetor[I];
    inc(Linha);

    // Algoritmo Heap's para gerar as permutações
    I := 0;
    while I < N do
    begin
      if Indices[I] < I then
      begin
        // Trocar elementos
        if I mod 2 = 0 then
          Temp := Vetor[0]
        else
          Temp := Vetor[Indices[I]];

        if I mod 2 = 0 then
          Vetor[0] := Vetor[I]
        else
          Vetor[Indices[I]] := Vetor[I];

        Vetor[I] := Temp;

        // Armazenar a nova permutação na matriz
        for j := 0 to N - 1 do
          Matriz[Linha][j] := Vetor[j];
        inc(Linha);

        // Incrementar o índice
        inc(Indices[I]);
        I := 0;
      end
      else
      begin
        Indices[I] := 0;
        inc(I);
      end;
    end;

    { if numcrit <= 7 then
      begin
      // Ordenar a matriz de permutações em ordem lexicográfica
      for I := 0 to Fatorial - 2 do
      for j := I + 1 to Fatorial - 1 do
      begin
      for k := 0 to N - 1 do
      begin
      if Matriz[I][k] < Matriz[j][k] then
      Break
      else if Matriz[I][k] > Matriz[j][k] then
      begin
      // Trocar as linhas
      for kk := 0 to N - 1 do
      begin
      Temp := Matriz[I][kk];
      Matriz[I][kk] := Matriz[j][kk];
      Matriz[j][kk] := Temp;
      end;
      Break;
      end;
      end;
      end;
      end; }

    SetLength(CasesOrdemcrit, Fatorial + 1, numcrit);
    SetLength(TiwUserSession(WebApplication.Data).NsolInst, Numalt);
    SetLength(TiwUserSession(WebApplication.Data).NsolInstPromethee, Numalt);
    for I := 0 to Fatorial - 1 do
      for j := 0 to numcrit - 1 do
        TiwUserSession(WebApplication.Data).CasesOrdemcrit[I, j] := Matriz
          [I, j];
  end;
end;

procedure TIWForm1.IWButton1Click(Sender: TObject);
begin
rgnalternativaY.Visible := true;
rgnalternativaX.Visible := False;
end;

procedure TIWForm1.IWButton2Click(Sender: TObject);
begin
rgnalternativaW.Visible := true;
rgnalternativaX.Visible := false;
end;

procedure TIWForm1.IWButton3Click(Sender: TObject);
begin
rgnalternativaZ.Visible := true;
rgnalternativaX.Visible := false;
end;

procedure TIWForm1.IWButton4Click(Sender: TObject);
begin
rgnalternativaS.Visible := true;
rgnalternativaX.Visible := false;
end;

procedure TIWForm1.IWLink10Click(Sender: TObject);
begin
rgnresultadosprobabilidade.Visible := true;
rgnalternativaX.Visible := false;
//rgnalternativaX.Align := AlClient;
end;

procedure TIWForm1.IWLink11Click(Sender: TObject);
begin
rgnresultadosprobabilidade.Visible := false;
rgnalternativaX.Visible := true;
rgnalternativaX.Align := AlClient;
end;

procedure TIWForm1.IWLink12Click(Sender: TObject);
begin
  rgnAnaliseEstatistica.Visible := false;
  rgnresultado.Visible := true;
end;

procedure TIWForm1.IWLink13Click(Sender: TObject);
begin
rgnalternativaX.Visible := true;
rgnperdamaxima.Visible := false;
end;

procedure TIWForm1.IWLink14Click(Sender: TObject);
begin
rgnalternativaX.Visible := true;
rgnalternativaZ.Visible := false;
end;

procedure TIWForm1.IWLink15Click(Sender: TObject);
begin
rgnalternativaX.Visible := true;
rgnalternativaY.Visible := false;
end;

procedure TIWForm1.IWLink16Click(Sender: TObject);
begin
rgnalternativaX.Visible := false;
rgnperdamaxima.Visible := true;
end;

procedure TIWForm1.IWLink17Click(Sender: TObject);
begin
rgnalternativaX.Visible := true;
rgnalternativaw.Visible := false;
end;

procedure TIWForm1.IWLink18Click(Sender: TObject);
begin
rgnalternativaY.visible := false;
rgnalternativaw.Visible := true;
end;

procedure TIWForm1.IWLink19Click(Sender: TObject);
begin
rgnalternativaS.Visible := False;
rgnalternativaX.visible := true;
end;

procedure TIWForm1.IWLink1Click(Sender: TObject);
begin
  rgnmatcons.Visible := true;
  rgnresultado.Visible := false;
end;

procedure TIWForm1.IWLink20Click(Sender: TObject);
begin
rgnperdamaxima.Visible := true;
rgnconsequencerange.Visible := false;
end;

procedure TIWForm1.IWLink21Click(Sender: TObject);
begin
rgnperdamaxima.Visible := false;
rgnconsequencerange.Visible := true;
end;

procedure TIWForm1.IWLink22Click(Sender: TObject);
begin
rgnconsequencerange.Visible := true;
rgndifvg.Visible := false;
end;

procedure TIWForm1.IWLink23Click(Sender: TObject);
begin
rgnconsequencerange.Visible := false;
rgndifvg.Visible := true;
end;

procedure TIWForm1.IWLink24Click(Sender: TObject);
begin
rgndifvg.Visible := true;
rgnmatrizconseqAltXS.Visible := false;
end;

procedure TIWForm1.IWLink25Click(Sender: TObject);
begin
rgndifvg.Visible := false;
rgnmatrizconseqAltXS.Visible := true;
end;

procedure TIWForm1.IWLink2Click(Sender: TObject);
begin
  // rgnestatistica.visible := false;
  /// rgnprimeiro.visible := true;
  // IWLink2.visible := false;
  /// lbltexto.Caption :=
  // 'Percentage of cases in which the top positions defined the solution';

  rgnresultado.Visible := false;
  rgnresultadosprobabilidade.Visible := true;
end;

procedure TIWForm1.IWLink3Click(Sender: TObject);
begin
  rgnestatistica.Visible := true;
  rgnprimeiro.Visible := false;
  IWLink2.Visible := true;
  lbltexto.Caption := 'Simulation Statistical Parameters';

end;

procedure TIWForm1.IWLink4Click(Sender: TObject);
begin
  rgnmatcons.Visible := true;
  rgnresultadoPromethee.Visible := false;
end;

procedure TIWForm1.IWLink6Click(Sender: TObject);
begin
  rgnregradecisaoPromethee.Visible := true;
  rgnestatisticapromethee.Visible := false;
end;

procedure TIWForm1.IWLink7Click(Sender: TObject);
begin
  // rgnregradecisao.Visible := true;
  // rgnestatistica.Visible := false;

  // rgnAnaliseEstatistica.Visible := true;
  // rgnAnaliseEstatistica.Align := alclient;
  rgnresultado.Visible := false;
  rgnregra.Visible := true;
end;

procedure TIWForm1.IWLink8Click(Sender: TObject);
begin
  rgnresultadosprobabilidade.Visible := false;
  rgnresultado.Visible := true;
end;

procedure TIWForm1.IWLink9Click(Sender: TObject);
begin
  rgnregra.Visible := false;
  rgnresultado.Visible := true;
end;

procedure TIWForm1.SetarCores;
Var
  I, j: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    SetLength(coresalt, 130);
    SetLength(coresalt1, TiwUserSession(WebApplication.Data).Numalt);
    coresalt[0] := $000000CC; // vermelho
    coresalt[1] := $0033BBFF; // amarelo
    coresalt[2] := $00CC9900; // azul
    coresalt[3] := $00BBBBFF; // rosa
    coresalt[4] := $005AE6C8; // verde
    coresalt[5] := $00FF00FF;
    coresalt[6] := $00808000;
    coresalt[7] := $00800000;
    coresalt[8] := $00000080;
    coresalt[9] := $0000FF00;
    coresalt[10] := $00008080;
    coresalt[11] := $00800080;
    coresalt[12] := $00FFFF00;
    coresalt[13] := $00C0DCC0;
    coresalt[14] := $00A4A0A0;
    coresalt[15] := $00E049BB;
    coresalt[16] := $005353FF;
    coresalt[17] := $000080FF;
    coresalt[18] := $00800040;
    coresalt[19] := $00434321;
    coresalt[20] := $00189AD8;
    coresalt[21] := $000000C1;
    coresalt[22] := $00DC0E65;
    coresalt[23] := $0000AE00;
    coresalt[24] := $002676CE;
    coresalt[25] := $008000FF; // índigo
    coresalt[26] := $00808080; // cinza claro
    coresalt[27] := $00FF4500; // laranja avermelhado
    coresalt[28] := $00FFD700; // ouro claro
    coresalt[29] := $00640000; // marrom escuro
    coresalt[30] := $008B4513; // marrom de sela
    coresalt[31] := $00A52A2A; // marrom de fogo
    coresalt[32] := $00D2691E; // chocolate
    coresalt[33] := $008BB000; // marrom escuro
    coresalt[34] := $00B22222; // tijolo
    coresalt[35] := $00FF6347; // tomate
    coresalt[36] := $00FF8C00; // laranja escuro
    coresalt[37] := $00FFA07A; // salmão claro
    coresalt[38] := $00FF7F50; // coral claro
    coresalt[39] := $00FF0000; // vermelho claro
    coresalt[40] := $00CD5C5C; // marrom claro
    coresalt[41] := $008B0000; // marrom escuro
    coresalt[42] := $00A52A2A; // marrom de fogo
    coresalt[43] := $00B22222; // tijolo
    coresalt[44] := $008B0000; // marrom escuro
    coresalt[45] := $00A52A2A; // marrom de fogo
    coresalt[46] := $00B22222; // tijolo
    coresalt[47] := $008B0000; // marrom escuro
    coresalt[48] := $00A52A2A; // marrom de fogo
    coresalt[49] := $00B22222; // tijolo
    coresalt[50] := $008B0000; // marrom escuro
    coresalt[51] := $00A52A2A; // marrom de fogo
    coresalt[52] := $00B22222; // tijolo
    coresalt[53] := $008B0000; // marrom escuro
    coresalt[54] := $00A52A2A; // marrom de fogo
    coresalt[55] := $00B22222; // tijolo
    coresalt[56] := $00FF69B4; // rosa quente
    coresalt[57] := $00FFE4E1; // rosa claro
    coresalt[58] := $00FF1493; // rosa profundo
    coresalt[59] := $00C71585; // vermelho violeta
    coresalt[60] := $00DB7093; // roxo claro
    coresalt[61] := $00DA70D6; // orquídea
    coresalt[62] := $009932CC; // índigo escuro
    coresalt[63] := $009ACD32; // verde amarelo
    coresalt[64] := $00C0FF3E; // verde limão brilhante
    coresalt[65] := $00AEEEEE; // azul celeste
    coresalt[66] := $000000CC; // vermelho
    coresalt[67] := $0033BBFF; // amarelo
    coresalt[68] := $00CC9900; // azul
    coresalt[69] := $00BBBBFF; // rosa
    coresalt[70] := $005AE6C8; // verde
    coresalt[71] := $00FF00FF;
    coresalt[72] := $00808000;
    coresalt[73] := $00800000;
    coresalt[74] := $00000080;
    coresalt[75] := $0000FF00;
    coresalt[76] := $00008080;
    coresalt[77] := $00800080;
    coresalt[78] := $00FFFF00;
    coresalt[79] := $00C0DCC0;
    coresalt[80] := $00A4A0A0;
    coresalt[81] := $00E049BB;
    coresalt[82] := $005353FF;
    coresalt[83] := $000080FF;
    coresalt[84] := $00800040;
    coresalt[85] := $00434321;
    coresalt[86] := $00189AD8;
    coresalt[87] := $000000C1;
    coresalt[88] := $00DC0E65;
    coresalt[89] := $0000AE00;
    coresalt[90] := $002676CE;
    coresalt[91] := $008000FF; // índigo
    coresalt[92] := $00808080; // cinza claro
    coresalt[93] := $00FF4500; // laranja avermelhado
    coresalt[94] := $00FFD700; // ouro claro
    coresalt[95] := $00640000; // marrom escuro
    coresalt[96] := $008B4513; // marrom de sela
    coresalt[97] := $00A52A2A; // marrom de fogo
    coresalt[98] := $00D2691E; // chocolate
    coresalt[99] := $008B0000; // marrom escuro
    coresalt[100] := $00B22222; // tijolo
    coresalt[101] := $00FF6347; // tomate
    coresalt[102] := $00FF8C00; // laranja escuro
    coresalt[103] := $00FFA07A; // salmão claro
    coresalt[104] := $00FF7F50; // coral claro
    coresalt[105] := $00FF0000; // vermelho claro
    coresalt[106] := $00CD5C5C; // marrom claro
    coresalt[107] := $008B0000; // marrom escuro
    coresalt[108] := $00A52A2A; // marrom de fogo
    coresalt[109] := $00B22222; // tijolo
    coresalt[110] := $008B0000; // marrom escuro
    coresalt[111] := $00A52A2A; // marrom de fogo
    coresalt[112] := $00B22222; // tijolo
    coresalt[113] := $008B0000; // marrom escuro
    coresalt[114] := $00A52A2A; // marrom de fogo
    coresalt[115] := $00B22222; // tijolo
    coresalt[116] := $008B0000; // marrom escuro
    coresalt[117] := $00A52A2A; // marrom de fogo
    coresalt[118] := $00B22222; // tijolo
    coresalt[119] := $008B0000; // marrom escuro
    coresalt[120] := $00A52A2A; // marrom de fogo
    coresalt[121] := $00B22222; // tijolo
    coresalt[122] := $00FF69B4; // rosa quente
    coresalt[123] := $00FFE4E1; // rosa claro
    coresalt[124] := $00FF1493; // rosa profundo
    coresalt[125] := $00C71585; // vermelho violeta
    coresalt[126] := $00DB7093; // roxo claro
    coresalt[127] := $00DA70D6; // orquídea
    coresalt[128] := $009932CC; // índigo escuro
    coresalt[129] := $009ACD32; // verde amarelo
  end;

end;

procedure TIWForm1.LimpaVariaveis;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    // Finalize(Tipocrit);
    // Finalize(niveis);
    // Finalize(MatrizConseq);
    // Finalize(Max);
    // Finalize(Min);
    // Finalize(MatrizConseqNorm);
    finalize(ResultSol);
    finalize(ResultAlt);
    // finalize(edtNomeCrit);
    // finalize(edtNomeAlt);
    // finalize(cmbbxNiveisCrit);
    // finalize(cmbbxTipoCrit);
    // finalize(edtMatrizConseq);
    // finalize(NomeCrit);
    // finalize(NomeAlt);
    // finalize(Tipocrit);
    // finalize(Niveis);
    // finalize(MatrizConseq);
    // finalize(MatrizConseqNorm);
    finalize(PesoCrit);
    finalize(CasesOrdemcrit);
    // finalize(Max);
    // finalize(Min);
    finalize(MatrizPOA);
    finalize(ResultadoROC);
    finalize(MatrizSol);
    finalize(ResultAlt);
    finalize(ResultSol);
    finalize(coresalt);
    finalize(coresalt1);
    finalize(CaseSol);
    finalize(CasesOrdemcrit);
    finalize(MatrizPOA);
    finalize(ResultadoROC);
    finalize(PesoCrit);
    // Finalize(NomeAlt);
    // Finalize(Nomecrit);
    finalize(MatrizSol);
    finalize(coresalt);
    finalize(coresalt1);
    finalize(NsolInst);
    finalize(Primcrit);
    finalize(Seccrit);
    finalize(terccrit);
    finalize(quartcrit);
    finalize(TotalPrimCrit);
    finalize(totalseccrit);
    finalize(totalterccrit);
    finalize(totalquartcrit);
    TiwUserSession(WebApplication.Data).percentualr1 := 0;
    TiwUserSession(WebApplication.Data).percentualr2 := 0;
    TiwUserSession(WebApplication.Data).percentualr3 := 0;
    TiwUserSession(WebApplication.Data).percentualr4 := 0;
    regra1 := 0;
    regra2 := 0;
    regra3 := 0;
    regra4 := 0;

    TiwUserSession(WebApplication.Data).percentualr1Promethee := 0;
    TiwUserSession(WebApplication.Data).percentualr2Promethee := 0;
    TiwUserSession(WebApplication.Data).percentualr3Promethee := 0;
    TiwUserSession(WebApplication.Data).percentualr4Promethee := 0;
    regra1Promethee := 0;
    regra2Promethee := 0;
    regra3Promethee := 0;
    regra4Promethee := 0;
  end;
end;

procedure TIWForm1.lnkMatrizConseqAltXClick(Sender: TObject);
var
i, j, k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
  k := cmbbxmatrizconseqAltXS.ItemIndex;
  if k = 0 then exit;


  grdMatrizConseqAltX.ClearCells;
  grdMatrizConseqAltX.Columns.Clear;

  for i := 0 to numcrit do
  grdMatrizConseqAltX.Columns.Add;

  grdmatrizconseqaltx.TotalRows := 9+numalt;

  grdmatrizconseqaltx.Cells[0,0] := 'Criteria:';
  grdmatrizconseqaltx.Cells[0,1] := 'Criterion type (FITradeoff):';
  grdmatrizconseqaltx.Cells[0,6] := 'Number of levels of discrete criteria:';
  grdmatrizconseqaltx.Cells[0,7] := 'Alternatives:';
  grdmatrizconseqaltx.Cells[1,7] := 'Consequence Matrix:';

  for i := 0 to numcrit - 1 do
  begin
   grdmatrizconseqaltx.Cells[1+i,0] := 'C' + inttostr(i+1);
   grdmatrizconseqaltx.Cells[1+i,1] := '1';
  end;

  for j := 0 to numalt - 1 do
  grdmatrizconseqaltx.Cells[0,8+j] := 'Alt ' + inttostr(j+1);


  for j := 0 to numalt - 1 do
  for i := 0 to numcrit - 1 do
  begin
   grdmatrizconseqaltx.Cells[1+i,8+j] := floattostr(MatrizConseqPercentilPerdaX[k*10-1,j,i]);
  end;


   tiwadvwebgridexcelio2.AdvWebGrid := grdmatrizconseqaltx;

 TIWAdvWebGridExcelIO2.XLSExport
        (TiwUserSession(WebApplication.Data).Nomeproblema + '   - Consequence Matrix for percentile = ' + floattostr(k/100) +'.xls', WebApplication);



  end;
end;

procedure TIWForm1.btnshowresultsClick(Sender: TObject);
var
  I, j, k, inst: Integer;
  bar, barpromethee: array of TBarSeries;
  Fatorial: Double;
  Naltsel: Integer;
  soma: Double;
  aux1, aux2: string;
  valor1, valor2: Integer;
  minIdx: Integer;
  tmp: Double;
  auxiliar1, auxiliar2, auxiliar3 : double;

begin


  With TiwUserSession(WebApplication.Data) do
  begin
    LimpaVariaveis;
    aditivo := false;
    sobreclassificacao := false;
    if chcklstbxMetodos.Selected[0] then
      aditivo := true;
    if chcklstbxMetodos.Selected[1] then
      sobreclassificacao := true;

    if (aditivo = false) and (sobreclassificacao = false) then
    begin
      WebApplication.ShowMessage('Select an model to continue!');
      Exit;
    end;

    aux1 := edtnumalt.Text;
    aux2 := cmbbxcriteria.Items[cmbbxcriteria.ItemIndex];
    if (TryStrToInt(aux1, valor1) = false) or
      ((TryStrToInt(aux2, valor2) = false)) then
    begin
      WebApplication.ShowMessage('Fill in the fields correctly!');
      Exit;
    end;

    NomeProblema := '';
    Nomeproblema := edtnumalt.Text + ' Alternatives and ' + cmbbxcriteria.Items[cmbbxcriteria.ItemIndex] + ' Criteria';
    lblproblema1.Caption := 'Simulation: ' + nomeproblema;
    lblproblema2.Caption := lblproblema1.Caption;
    lblproblema3.Caption := lblproblema2.Caption;
    lblproblema4.Caption := lblproblema3.Caption;
    lblproblema5.Caption := lblproblema4.Caption;
    lblproblema6.Caption := lblproblema5.Caption;
    lblproblema7.Caption := lblproblema6.Caption;
    lblproblema8.Caption := lblproblema7.Caption;
    lblproblema9.caption := lblproblema8.Caption;
    Numalt := StrToInt(edtnumalt.Text);
    numcrit := StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    SetLength(Nomecrit, numcrit);
    SetLength(Nomealt, Numalt);
    SetLength(tipocrit, numcrit);
    SetLength(niveis, numcrit);
    SetLength(MatrizConseq, Numalt, numcrit);
    SetLength(MatrizConseqNorm, Numalt, numcrit);
    SetLength(Max, numcrit);
    SetLength(Min, numcrit);
    GerarCases;
    SetLength(PesoCrit, Length(CasesOrdemcrit), numcrit);
    Ninst := 1000;
    SetLength(perdamediapopulacao, Ninst);
    SetLength(perdamaximapopulacao, Ninst);
    SetLength(problimiarperdamedia, 1000);
    SetLength(problimiarperdamaxima, 1000);
    SetLength(perdamediaindividuos, Numalt);
    SetLength(ganhomedioindividuos, Numalt);
    SetLength(desempenhoindividuos, Numalt);
    SetLength(perdamaximaindividuos, Numalt);
    SetLength(desempenhoindividuos, Numalt);
    SetLength(desempenhopopulacao, Ninst);
    SetLength(matrizprobalt, Ninst);
    setlength(MatrizFaixaPrimeira, numalt, 6);
    setlength(MatrizFaixaSegunda, numalt, 6);
    setlength(MatrizFaixaTerceira, numalt, 6);
    setlength(VetorAlternativaX, ninst, 5);
    setlength(VetorAlternativaZ, ninst, 5);
    setlength(VetorAlternativaS, ninst, 5);

    setlength(QuartisAlternativaX, 5, 4);
    setlength(QuartisAlternativaZ, 5, 4);
    setlength(QuartisAlternativaS, 5, 4);

    setlength(PercentilAlternativaX, 5, 101);
    setlength(PercentilAlternativaZ, 5, 101);
    setlength(PercentilAlternativaS, 5, 101);

    setlength(vetoralternativaY, ninst, 2);
    setlength(quartisalternativaY, 2, 4);
    setlength(percentilalternativaY, 2, 101);

    setlength(vetorAlternativaW, ninst, 2);
    setlength(quartisalternativaW, 2, 4);
    setlength(percentilalternativaW, 2, 101);

    setlength(TiwUserSession(WebApplication.Data).PerdaMaximaX, ninst);
    setlength(TiwUserSession(WebApplication.Data).PerdaMaximaZ, ninst);
    setlength(TiwUserSession(WebApplication.Data).PerdaMaximaS, ninst);
    setlength(TiwUserSession(WebApplication.Data).PerdaMaximaW, ninst);

    setlength(TiwUserSession(WebApplication.Data).OrdemPerdaMaximaX, ninst);
    setlength(TiwUserSession(WebApplication.Data).OrdemPerdaMaximaZ, ninst);
    setlength(TiwUserSession(WebApplication.Data).OrdemPerdaMaximaS, ninst);
    setlength(TiwUserSession(WebApplication.Data).OrdemPerdaMaximaW, ninst);

    setlength(QuartisPerdaMaximaX, 4);
    setlength(QuartisPerdaMaximaZ, 4);
    setlength(QuartisPerdaMaximaS, 4);
    setlength(QuartisPerdaMaximaW, 4);

    setlength(PercentilPerdaMaximaX, 101);
    setlength(PercentilPerdaMaximaZ, 101);
    setlength(PercentilPerdaMaximaS, 101);
    setlength(PercentilPerdaMaximaW, 101);



    setlength(TiwUserSession(WebApplication.Data).ConsequenceRangeX, ninst);
    setlength(TiwUserSession(WebApplication.Data).ConsequenceRangeS, ninst);
    setlength(TiwUserSession(WebApplication.Data).ConsequenceRangez, ninst);

    setlength(quartisConsequenceRangeX, 4);
    setlength(quartisConsequenceRanges, 4);
    setlength(quartisConsequenceRangez, 4);

    setlength(PercentilConsequenceRangex, 101);
    setlength(PercentilConsequenceRanges, 101);
    setlength(PercentilConsequenceRangez, 101);


    setlength(TiwUserSession(WebApplication.Data).difVGX, ninst);
    setlength(TiwUserSession(WebApplication.Data).difVGZ, ninst);
    setlength(TiwUserSession(WebApplication.Data).difVGS, ninst);

    setlength(quartisdifVGX, 4);
    setlength(quartisdifVGZ, 4);
    setlength(quartisdifVGS, 4);

    setlength(PercentildifVGX, 101);
    setlength(PercentildifVGZ, 101);
    setlength(PercentildifVGS, 101);

    setlength(matrizconseqAltX, ninst, numcrit);
    setlength(matrizconseqAltS, ninst, numcrit);


    setlength(matrizconseqAltXOrdenado, ninst, numcrit);
    setlength(matrizconseqAltSOrdenado, ninst, numcrit);

    setlength(MatrizConseqPercentilPerdaX, ninst, numalt, numcrit);
    //setlength(MatrizConseqPercentilPerdaS, ninst, numalt, numcrit);

    mediadesempenho := 0;
    maximodesempenho := 0;

    for j := 0 to Numalt - 1 do
    begin
      Nomealt[j] := 'Alt ' + inttostr(j + 1);
    end;
    for I := 0 to numcrit - 1 do
    begin
      TiwUserSession(WebApplication.Data).tipocrit[I] := 1;
      Nomecrit[I] := 'Crit ' + inttostr(I + 1);
    end;
    for k := 0 to Length(CasesOrdemcrit) - 2 do
    begin
      for I := 0 to Length(CasesOrdemcrit[0]) - 1 do
      begin
        soma := 0;
        for j := I + 1 to TiwUserSession(WebApplication.Data).numcrit do
        begin
          soma := soma + 1 / j;
        end;
        TiwUserSession(WebApplication.Data)
          .PesoCrit[k, CasesOrdemcrit[k, I] - 1] :=
          (1 / TiwUserSession(WebApplication.Data).numcrit) * soma;
        // simpeso[indexroc[i]] := auxpeso
      end;
    end;
    k := Length(CasesOrdemcrit) - 1;
    for I := 0 to numcrit - 1 do
      PesoCrit[k, I] := 1 / numcrit;

    for inst := 0 to Ninst - 1 do
    begin
      for j := 0 to Numalt - 1 do
        for I := 0 to numcrit - 1 do
        begin
          TiwUserSession(WebApplication.Data).MatrizConseq[j, I] := Random;
          //MatrizConseqNorm[j, I] := MatrizConseq[j, I];
        end;

       for i := 0 to numcrit - 1 do
        begin
        if (Tipocrit[i] <> 2) and (Tipocrit[i] <> 3) then
        begin
        Max[i] := -9999999999;
        Min[i] := 9999999999;
        for j := 0 to Numalt - 1 do
        begin
        if Max[i] < MatrizConseq[j, i] then
        Max[i] := MatrizConseq[j, i];
        if Min[i] > MatrizConseq[j, i] then
        Min[i] := MatrizConseq[j, i];
        end; // j
        end; // if
        end; // i
        for j := 0 to Numalt - 1 do
        for i := 0 to numcrit - 1 do
        begin
        if (Tipocrit[i] = 1) or (Tipocrit[i] = 3) or (Tipocrit[i] = 5) then
        MatrizConseqNorm[j, i] := (MatrizConseq[j, i] - Min[i]) /
        (Max[i] - Min[i]);
        /// Normalização para os critérios de minimização //////////////////////
        if (Tipocrit[i] = 0) or (Tipocrit[i] = 2) or (Tipocrit[i] = 4) then
        TiwUserSession(WebApplication.Data).MatrizConseqNorm[j, i] :=
        (MatrizConseq[j, i] - Max[i]) / (-Max[i] + Min[i]);
        end;

      instancia := inst;
      Surrogate;
      if aditivo then
      begin
        Naltsel := 0;
        for j := 0 to Numalt - 1 do
        begin
          if probalt[j] > 0 then
            inc(Naltsel);
        end;
        for j := 0 to Numalt - 1 do
        begin
          if probalt[j] > 0 then
          begin
            perdamediapopulacao[inst] := perdamediapopulacao[inst]
              + perdamediaindividuos[j];
            // if perdamaximapopulacao[inst] < perdamaximaindividuos[j] then
            perdamaximapopulacao[inst] := perdamaximaindividuos[j]
              + perdamaximapopulacao[inst];
          end;
        end;
        TiwUserSession(WebApplication.Data).perdamediapopulacao[inst] :=
          perdamediapopulacao[inst] / Naltsel;
        TiwUserSession(WebApplication.Data).perdamaximapopulacao[inst] :=
          perdamaximapopulacao[inst] / Naltsel;

        for I := 0 to 1000 - 1 do
        begin
          if (I / 1000) <= perdamediapopulacao[inst] then
            TiwUserSession(WebApplication.Data).problimiarperdamedia[I] :=
              problimiarperdamedia[I] + 1;

          if (I / 1000) <= perdamaximapopulacao[inst] then
            TiwUserSession(WebApplication.Data).problimiarperdamaxima[I] :=
              problimiarperdamaxima[I] + 1;
        end;


      inc(TiwUserSession(WebApplication.Data).NsolInst[Length(TiwUserSession(WebApplication.Data).ResultSol) - 1]
          );
      end;
      if sobreclassificacao then
        inc(TiwUserSession(WebApplication.Data).NsolInstPromethee
            [Length(TiwUserSession(WebApplication.Data).ResultSolPromethee) - 1]
          );
    end;
    if aditivo then
    begin



      TiwUserSession(WebApplication.Data).PercentualPerdaMaximaXPesosIguais := PercentualPerdaMaximaXPesosIguais / ninst;
      TiwUserSession(WebApplication.Data).PercentualPerdaMaximaZPesosIguais := PercentualPerdaMaximaZPesosIguais / ninst;
      TiwUserSession(WebApplication.Data).PercentualPerdaMaximaSPesosIguais := PercentualPerdaMaximaSPesosIguais / ninst;

      OrdenarColunas;
      QuartisPorColuna;

      for j := 0 to Numalt - 1 do
      begin
        TiwUserSession(WebApplication.Data).RankingPrimeira[j] :=
          RankingPrimeira[j] / (Length(CasesOrdemcrit) * Ninst);
       // TiwUserSession(WebApplication.Data).RankingSegunda[j] := RankingSegunda
        //  [j] / (Length(CasesOrdemcrit) * Ninst);
       // TiwUserSession(WebApplication.Data).RankingTerceira[j] :=
        //  RankingTerceira[j] / (Length(CasesOrdemcrit) * Ninst);

      for I := 0 to 5 do
      begin
        TiwUserSession(WebApplication.Data).MatrizFaixaPrimeira[j,i] := MatrizFaixaPrimeira[j,i] / ninst;
        //TiwUserSession(WebApplication.Data).MatrizFaixaSegunda[j,i] := MatrizFaixaSegunda[j,i] / ninst;
        //TiwUserSession(WebApplication.Data).MatrizFaixaTerceira[j,i] := MatrizFaixaTerceira[j,i] / ninst;
      end;

      end;
      TiwUserSession(WebApplication.Data).mediadesempenho :=
        mediadesempenho / Ninst;

      for I := 0 to 100 do
      begin
        TiwUserSession(WebApplication.Data).probaltestratificado[I] :=
          100 * probaltestratificado[I] / Ninst;
        TiwUserSession(WebApplication.Data).probaltestratificado2[I] :=
          100 * probaltestratificado2[I] / Ninst;
        TiwUserSession(WebApplication.Data).probaltestratificado3[I] :=
          100 * probaltestratificado3[I] / Ninst;
      end;

      for I := 0 to 200 do
        percentualdesempenho[I] := 100 * percentualdesempenho[I] / Ninst;

      // supondo que matrizprobalt já foi declarado como: array of Double

      TiwUserSession(WebApplication.Data).mediaprobalt :=
        100 * mediaprobalt / Ninst;

      TiwUserSession(WebApplication.Data).mediaprobalt2 :=
        100 * mediaprobalt2 / Ninst;

      TiwUserSession(WebApplication.Data).mediaprobalt3 :=
        100 * mediaprobalt3 / Ninst;

      TiwUserSession(WebApplication.Data).maximoprobalt := 100 * maximoprobalt;

      DesvioPadraoprobalt := 0;
      for I := 0 to Length(probaltestratificado) - 1 do
        if probaltestratificado[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadraoprobalt :=
            DesvioPadraoprobalt + probaltestratificado[I] * Sqr
            (abs((I + 1) - mediaprobalt))
        end;
      DesvioPadraoprobalt := Sqrt(DesvioPadraoprobalt / Length
          (probaltestratificado));

      DesvioPadraoprobalt2 := 0;
      for I := 0 to Length(probaltestratificado2) - 1 do
        if probaltestratificado2[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadraoprobalt2 :=
            DesvioPadraoprobalt2 + probaltestratificado2[I] * Sqr
            (abs((I + 1) - mediaprobalt2))
        end;
      DesvioPadraoprobalt2 := Sqrt(DesvioPadraoprobalt2 / Length
          (probaltestratificado2));

      DesvioPadraoprobalt3 := 0;
      for I := 0 to Length(probaltestratificado3) - 1 do
        if probaltestratificado3[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadraoprobalt3 :=
            DesvioPadraoprobalt3 + probaltestratificado3[I] * Sqr
            (abs((I + 1) - mediaprobalt3))
        end;
      DesvioPadraoprobalt3 := Sqrt(DesvioPadraoprobalt3 / Length
          (probaltestratificado3));

      { mediana := 0;
        j := 0;
        for I := 0 to length(probaltestratificado) - 1 do
        if probaltestratificado[I] <> 0 then
        begin
        j := j + Round(probaltestratificado[I]);
        if j = Round(Ninst / 2) then
        begin
        medianaprobalt := ((I + 1) + (I + 2)) / 2;
        Break;
        end;
        if j > Round(Ninst / 2) then
        begin
        medianaprobalt := I + 1;
        Break;
        end;
        end;
        mediana := 100 * mediana;

        modaprobalt := 0;
        j := 0;
        for I := 0 to Length(probaltestratificado) - 1 do
        if probaltestratificado[I] <> 0 then
        begin
        if j < probaltestratificado[I] then
        begin
        j := Round(probaltestratificado[I]);
        modaprobalt := I + 1;
        end;
        end; }

      for I := 40 to 49 do
      begin
        TiwUserSession(WebApplication.Data).faixa40probalt :=
          faixa40probalt + TiwUserSession(WebApplication.Data)
          .probaltestratificado[I];
        TiwUserSession(WebApplication.Data).faixa40probalt2 :=
          faixa40probalt2 + TiwUserSession(WebApplication.Data)
          .probaltestratificado2[I];
        TiwUserSession(WebApplication.Data).faixa40probalt3 :=
          faixa40probalt3 + probaltestratificado3[I];
      end;

      for I := 50 to 59 do
      begin
        TiwUserSession(WebApplication.Data).faixa50probalt :=
          faixa50probalt + probaltestratificado[I];
        TiwUserSession(WebApplication.Data).faixa50probalt2 :=
          faixa50probalt2 + probaltestratificado2[I];
        TiwUserSession(WebApplication.Data).faixa50probalt3 :=
          faixa50probalt3 + probaltestratificado3[I];
      end;

      for I := 60 to 69 do
      begin
        TiwUserSession(WebApplication.Data).faixa60probalt :=
          faixa60probalt + probaltestratificado[I];
        TiwUserSession(WebApplication.Data).faixa60probalt2 :=
          faixa60probalt2 + probaltestratificado2[I];
        TiwUserSession(WebApplication.Data).faixa60probalt3 :=
          faixa60probalt3 + probaltestratificado3[I];
      end;

      for I := 70 to 79 do
      begin
        TiwUserSession(WebApplication.Data).faixa70probalt :=
          faixa70probalt + probaltestratificado[I];
        TiwUserSession(WebApplication.Data).faixa70probalt2 :=
          faixa70probalt2 + probaltestratificado2[I];
        TiwUserSession(WebApplication.Data).faixa70probalt3 :=
          faixa70probalt3 + probaltestratificado3[I];
      end;

      for I := 80 to 89 do
      begin
        TiwUserSession(WebApplication.Data).faixa80probalt := faixa80probalt + probaltestratificado[I];
        faixa80probalt2 := faixa80probalt2 + probaltestratificado2[I];
        faixa80probalt3 := faixa80probalt3 + probaltestratificado3[I];
      end;

      for I := 90 to 100 do
      begin
        TiwUserSession(WebApplication.Data).faixa90probalt := faixa90probalt + probaltestratificado[I];
        faixa90probalt2 := faixa90probalt2 + probaltestratificado2[I];
        faixa90probalt3 := faixa90probalt3 + probaltestratificado3[I];
      end;

      TiwUserSession(WebApplication.Data).faixa40probaltacumulado := faixa40probalt + faixa50probalt +
        faixa60probalt + faixa70probalt + faixa80probalt + faixa90probalt;

      faixa50probaltacumulado := faixa50probalt + faixa60probalt +
        faixa70probalt + faixa80probalt + faixa90probalt;

      faixa60probaltacumulado := faixa60probalt + faixa70probalt +
        faixa80probalt + faixa90probalt;

      faixa70probaltacumulado := faixa70probalt + faixa80probalt +
        faixa90probalt;

      faixa80probaltacumulado := faixa80probalt + faixa90probalt;

      TiwUserSession(WebApplication.Data).faixa90probaltacumulado :=
        faixa90probalt;
      // prob acumulada para 1ª e 2ª alternativa
      TiwUserSession(WebApplication.Data).faixa40probaltacumulado2 :=
        faixa40probalt2 + faixa50probalt2 + faixa60probalt2 + faixa70probalt2 +
        faixa80probalt2 + faixa90probalt2;

      TiwUserSession(WebApplication.Data).faixa50probaltacumulado2 :=
        faixa50probalt2 + faixa60probalt2 + faixa70probalt2 + faixa80probalt2 +
        faixa90probalt2;

      TiwUserSession(WebApplication.Data).faixa60probaltacumulado2 :=
        faixa60probalt2 + faixa70probalt2 + faixa80probalt2 + faixa90probalt2;

      TiwUserSession(WebApplication.Data).faixa70probaltacumulado2 :=
        faixa70probalt2 + faixa80probalt2 + faixa90probalt2;

      TiwUserSession(WebApplication.Data).faixa80probaltacumulado2 :=
        faixa80probalt2 + faixa90probalt2;

      TiwUserSession(WebApplication.Data).faixa90probaltacumulado2 :=
        faixa90probalt2;

      // prob acumulada para 1ª, 2ª e 3ª alternativa
      TiwUserSession(WebApplication.Data).faixa40probaltacumulado3 :=
        faixa40probalt3 + faixa50probalt3 + faixa60probalt3 + faixa70probalt3 +
        faixa80probalt3 + faixa90probalt3;

      TiwUserSession(WebApplication.Data).faixa50probaltacumulado3 :=
        faixa50probalt3 + faixa60probalt3 + faixa70probalt3 + faixa80probalt3 +
        faixa90probalt3;

      TiwUserSession(WebApplication.Data).faixa60probaltacumulado3 :=
        faixa60probalt3 + faixa70probalt3 + faixa80probalt3 + faixa90probalt3;

      TiwUserSession(WebApplication.Data).faixa70probaltacumulado3 :=
        faixa70probalt3 + faixa80probalt3 + faixa90probalt3;

      TiwUserSession(WebApplication.Data).faixa80probaltacumulado3 :=
        faixa80probalt3 + faixa90probalt3;

      TiwUserSession(WebApplication.Data).faixa90probaltacumulado3 :=
        faixa90probalt3;

      chrtprobalt1.series[0].Clear;
      chrtprobalt2.series[0].Clear;
      chrtprobalt3.series[0].Clear;
      for I := 0 to 10 do
      begin
        chrtprobalt1.BottomAxis.Items.Add(I * 10, inttostr(I * 10) + ' %');

        chrtprobalt2.BottomAxis.Items.Add(I * 10, inttostr(I * 10) + ' %');

        chrtprobalt3.BottomAxis.Items.Add(I * 10, inttostr(I * 10) + ' %');
      end;

      for I := 0 to 100 do
      begin
        chrtprobalt1.series[0].AddXY(I, probaltestratificado[I], formatfloat
            ('0', probaltestratificado[I]) + ' %');

        chrtprobalt2.series[0].AddXY(I, probaltestratificado2[I], formatfloat
            ('0', probaltestratificado2[I]) + ' %');

        chrtprobalt3.series[0].AddXY(I, probaltestratificado3[I], formatfloat
            ('0', probaltestratificado3[I]) + ' %');
        // chrtprobalt1.BottomAxis.items.Add(i, inttostr(i*10) + ' %');
      end;
      imgresultadosprobabilidade.Picture.Bitmap.Assign
        (chrtprobalt1.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
            imgresultadosprobabilidade.Width, imgresultadosprobabilidade.Height)
          ));

   {   chrtregra.series[0].Clear;
      for I := 0 to 20 do
        chrtregra.BottomAxis.Items.Add(I * 10, floattostr(I / 10 - 1));

      for I := 0 to 200 do
      begin
        percentualdesempenho[I] := 100 * percentualdesempenho[I] / Ninst;
        chrtregra.series[0].AddXY(I, percentualdesempenho[I], formatfloat
            ('0', percentualdesempenho[I]) + ' %');
        // chrtprobalt1.BottomAxis.items.Add(i, inttostr(i*10) + ' %');
      end;
      imgregra.Picture.Bitmap.Assign
        (chrtregra.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgregra.Width,
            imgregra.Height)));

      DesvioPadrao := 0.0;
      if Length(desempenhopopulacao) > 0 then
      begin
        for I := 0 to High(desempenhopopulacao) do
        begin
          // soma dos quadrados das diferenças entre cada valor e a medida central
          DesvioPadrao := DesvioPadrao + Sqr
            (desempenhopopulacao[I] - mediadesempenho);
        end;

        // Divide por N (populacional) e tira raiz
        DesvioPadrao := Sqrt(DesvioPadrao / Length(desempenhopopulacao));
      end
      else
        DesvioPadrao := 0.0;

      // --- mediana (valor Double, mesma lógica de acumular Round(...), sem variáveis novas) ---
      mediana := 0.0;
      j := 0;
      if Length(desempenhopopulacao) > 0 then
      begin
        for I := 0 to High(desempenhopopulacao) do
        begin
          j := j + round(desempenhopopulacao[I]);

          if j = round(Ninst / 2.0) then
          begin
            // média dos dois valores vizinhos (I e I+1) — não usa índices, usa os valores reais
            if I < High(desempenhopopulacao) then
              mediana := (desempenhopopulacao[I] + desempenhopopulacao[I + 1])
                / 2.0
            else
              mediana := desempenhopopulacao[I];
            Break;
          end;

          if j > round(Ninst / 2.0) then
          begin
            mediana := desempenhopopulacao[I];
            Break;
          end;
        end;

        // fallback: se nunca cobriu metade, usa o último valor do array
        if mediana = 0.0 then
          mediana := desempenhopopulacao[ High(desempenhopopulacao)];
      end
      else
        mediana := 0.0;

      // --- Moda (valor Double, mesma lógica: maior Round(desempenhopopulacao[I])) ---
      moda := 0.0;
      j := 0;
      if Length(percentualdesempenho) > 0 then
      begin
        for I := 0 to High(percentualdesempenho) do
          if percentualdesempenho[I] <> 0 then
          begin
            if j < round(100 * percentualdesempenho[I]) then
            begin
              j := round(100 * percentualdesempenho[I]);
              moda := (I - 100) / 100;
            end;
          end;
      end;           }


      chrtalternativaX.Series[0].Clear;
      chrtalternativaX.Series[1].Clear;
      chrtalternativaX.Series[2].Clear;
      chrtalternativaX.Series[3].Clear;
      chrtalternativaX.Series[4].Clear;

      chrtalternativaX.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaX.Series[1].Marks.FontSeriesColor := true;
      chrtalternativaX.Series[2].Marks.FontSeriesColor := true;
      chrtalternativaX.Series[3].Marks.FontSeriesColor := true;
      chrtalternativaX.Series[4].Marks.FontSeriesColor := true;

      chrtalternativaX.BottomAxis.Items.Clear;
      chrtalternativaX.BottomAxis.Items.Add(0,'0.25');
      chrtalternativaX.BottomAxis.Items.Add(1,'0.50');
      chrtalternativaX.BottomAxis.Items.Add(2,'0.75');
      chrtalternativaX.BottomAxis.Items.Add(3,'1.00');
      for j := 0 to 3 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaX.Series[0].AddXY(j, QuartisAlternativaX[0,j]*100, formatfloat('0', QuartisAlternativaX[0,j]*100) + '%');
          chrtalternativaX.Series[1].AddXY(j, QuartisAlternativaX[1,j]*100, formatfloat('0', QuartisAlternativaX[1,j]*100) + '%');
          chrtalternativaX.Series[2].AddXY(j, QuartisAlternativaX[2,j]*100, formatfloat('0', QuartisAlternativaX[2,j]*100) + '%');
          chrtalternativaX.Series[3].AddXY(j, QuartisAlternativaX[3,j]*100, formatfloat('0', QuartisAlternativaX[3,j]*100) + '%');
          chrtalternativaX.Series[4].AddXY(j, QuartisAlternativaX[4,j]*100, formatfloat('0', QuartisAlternativaX[4,j]*100) + '%');
          //chrtalternativaX.Series[1].AddXY(j, rankingsegunda[j]*100, formatfloat('0', rankingsegunda[j]*100) + '%');
          //chrtalternativaX.Series[2].AddXY(j, rankingterceira[j]*100, formatfloat('0', rankingterceira[j]*100) + '%');

        //end;
      end;


      chrtalternativaX2.Series[0].Clear;
      chrtalternativaX2.Series[1].Clear;
      chrtalternativaX2.Series[2].Clear;
      chrtalternativaX2.Series[3].Clear;
      chrtalternativaX2.Series[4].Clear;

      chrtalternativaX2.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaX2.Series[1].Marks.FontSeriesColor := true;
      chrtalternativaX2.Series[2].Marks.FontSeriesColor := true;
      chrtalternativaX2.Series[3].Marks.FontSeriesColor := true;
      chrtalternativaX2.Series[4].Marks.FontSeriesColor := true;

      chrtalternativaX2.BottomAxis.Items.Clear;
      chrtalternativaX2.BottomAxis.Items.Add(25,'0.25');
      chrtalternativaX2.BottomAxis.Items.Add(50,'0.50');
      chrtalternativaX2.BottomAxis.Items.Add(75,'0.75');
      chrtalternativaX2.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaX2.Series[0].AddXY(j, PercentilAlternativaX[0,j]*100, formatfloat('0', PercentilAlternativaX[0,j]*100) + '%');
          chrtalternativaX2.Series[1].AddXY(j, PercentilAlternativaX[1,j]*100, formatfloat('0', PercentilAlternativaX[1,j]*100) + '%');
          chrtalternativaX2.Series[2].AddXY(j, PercentilAlternativaX[2,j]*100, formatfloat('0', PercentilAlternativaX[2,j]*100) + '%');
          chrtalternativaX2.Series[3].AddXY(j, PercentilAlternativaX[3,j]*100, formatfloat('0', PercentilAlternativaX[3,j]*100) + '%');
          chrtalternativaX2.Series[4].AddXY(j, PercentilAlternativaX[4,j]*100, formatfloat('0', PercentilAlternativaX[4,j]*100) + '%');
          //chrtalternativaY.Series[1].AddXY(j, rankingsegunda[j]*100, formatfloat('0', rankingsegunda[j]*100) + '%');
          //chrtalternativaY.Series[2].AddXY(j, rankingterceira[j]*100, formatfloat('0', rankingterceira[j]*100) + '%');

        //end;
      end;
      imgalternativaX.Picture.Bitmap.Assign
        (chrtalternativaX2.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgalternativaX.Width,
            imgalternativaX.Height)));

      grdalternativaX.Clear;
      grdalternativaX.ColumnCount := 5;
      grdalternativaX.RowCount := 6;

      for j := 0 to grdalternativaX.ColumnCount - 1 do
        for I := 0 to grdalternativaX.RowCount - 1 do
        begin
          grdalternativaX.Cell[I, j].Alignment := taCenter;
          grdalternativaX.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdalternativaX.ColumnCount - 1 do
        grdalternativaX.Cell[0, j].Font.Style := [fsbold];

      grdalternativaX.Cell[0, 0].Text := 'Position';
      grdalternativaX.Cell[1, 0].Text := '1º';
      grdalternativaX.Cell[2, 0].Text := '2º';
      grdalternativaX.Cell[3, 0].Text := '3º';
      grdalternativaX.Cell[4, 0].Text := '1º or 2º';
      grdalternativaX.Cell[5, 0].Text := '1º or 2º or 3º';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdalternativaX.Cell[0, 1].Text := 'Percentile = 0.25';
      grdalternativaX.Cell[0, 2].Text := 'Percentile = 0.50';
      grdalternativaX.Cell[0, 3].Text := 'Percentile = 0.75';
      grdalternativaX.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      for j := 0 to 4 do
      begin
        grdalternativaX.Cell[j+1, i+1].Text := formatfloat('0', QuartisAlternativaX[j,i] * 100) + '%';
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;

      chrtalternativaS.Series[0].Clear;
      chrtalternativaS.Series[1].Clear;
      chrtalternativaS.Series[2].Clear;
      chrtalternativaS.Series[3].Clear;
      chrtalternativaS.Series[4].Clear;
      //chrtalternativaS.Series[5].Clear;

      chrtalternativaS.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaS.Series[1].Marks.FontSeriesColor := true;
      chrtalternativaS.Series[2].Marks.FontSeriesColor := true;
      chrtalternativaS.Series[3].Marks.FontSeriesColor := true;
      chrtalternativaS.Series[4].Marks.FontSeriesColor := true;
      //chrtalternativaS.Series[5].Marks.FontSeriesColor := true;

      chrtalternativaS.BottomAxis.Items.Clear;
      chrtalternativaS.BottomAxis.Items.Add(25,'0.25');
      chrtalternativaS.BottomAxis.Items.Add(50,'0.50');
      chrtalternativaS.BottomAxis.Items.Add(75,'0.75');
      chrtalternativaS.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaS.Series[0].AddXY(j, PercentilAlternativaS[0,j]*100, formatfloat('0', PercentilAlternativaS[0,j]*100) + '%');
          chrtalternativaS.Series[1].AddXY(j, PercentilAlternativaS[1,j]*100, formatfloat('0', PercentilAlternativaS[1,j]*100) + '%');
          chrtalternativaS.Series[2].AddXY(j, PercentilAlternativaS[2,j]*100, formatfloat('0', PercentilAlternativaS[2,j]*100) + '%');
          chrtalternativaS.Series[3].AddXY(j, PercentilAlternativaS[3,j]*100, formatfloat('0', PercentilAlternativaS[3,j]*100) + '%');
          chrtalternativaS.Series[4].AddXY(j, PercentilAlternativaS[4,j]*100, formatfloat('0', PercentilAlternativaS[4,j]*100) + '%');
          //chrtalternativaS.Series[5].AddXY(j, PercentilAlternativaS[5,j]*100, formatfloat('0', PercentilAlternativaS[5,j]*100) + '%');
          //chrtalternativaY.Series[1].AddXY(j, rankingsegunda[j]*100, formatfloat('0', rankingsegunda[j]*100) + '%');
          //chrtalternativaY.Series[2].AddXY(j, rankingterceira[j]*100, formatfloat('0', rankingterceira[j]*100) + '%');

        //end;
      end;
      imgalternativaS.Picture.Bitmap.Assign
        (chrtalternativaS.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgalternativaS.Width,
            imgalternativaS.Height)));

      grdalternativaS.Clear;
      grdalternativaS.ColumnCount := 5;
      grdalternativaS.RowCount := 6;

      for j := 0 to grdalternativaS.ColumnCount - 1 do
        for I := 0 to grdalternativaS.RowCount - 1 do
        begin
          grdalternativaS.Cell[I, j].Alignment := taCenter;
          grdalternativaS.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdalternativaS.ColumnCount - 1 do
        grdalternativaS.Cell[0, j].Font.Style := [fsbold];

      grdalternativaS.Cell[0, 0].Text := 'Position';
      grdalternativaS.Cell[1, 0].Text := '1º';
      grdalternativaS.Cell[2, 0].Text := '2º';
      grdalternativaS.Cell[3, 0].Text := '3º';
      grdalternativaS.Cell[4, 0].Text := '1º or 2º';
      grdalternativaS.Cell[5, 0].Text := '1º or 2º or 3º';
      //grdalternativaS.Cell[6, 0].Text := 'S = X';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdalternativaS.Cell[0, 1].Text := 'Percentile = 0.25';
      grdalternativaS.Cell[0, 2].Text := 'Percentile = 0.50';
      grdalternativaS.Cell[0, 3].Text := 'Percentile = 0.75';
      grdalternativaS.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      for j := 0 to 4 do
      begin
        grdalternativaS.Cell[j+1, i+1].Text := formatfloat('0', QuartisAlternativaS[j,i] * 100) + '%';
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';
      end;
      PercentualSigualaX := PercentualSigualaX / ninst;
      lblPercentualSigualaX.Caption := 'The alternative S is equal to alternative X in ' + formatfloat('0', PercentualSigualaX * 100) + '% of instances';

      chrtalternativaZ.Series[0].Clear;
      chrtalternativaZ.Series[1].Clear;
      chrtalternativaZ.Series[2].Clear;
      chrtalternativaZ.Series[3].Clear;
      chrtalternativaZ.Series[4].Clear;

      chrtalternativaZ.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaZ.Series[1].Marks.FontSeriesColor := true;
      chrtalternativaZ.Series[2].Marks.FontSeriesColor := true;
      chrtalternativaZ.Series[3].Marks.FontSeriesColor := true;
      chrtalternativaZ.Series[4].Marks.FontSeriesColor := true;

      chrtalternativaZ.BottomAxis.Items.Clear;
      chrtalternativaZ.BottomAxis.Items.Add(0,'0.25');
      chrtalternativaZ.BottomAxis.Items.Add(1,'0.50');
      chrtalternativaZ.BottomAxis.Items.Add(2,'0.75');
      chrtalternativaZ.BottomAxis.Items.Add(3,'1.00');
      for j := 0 to 3 do
      begin
        //if (rankingsegunda[j] <> 0) and (round(rankingsegunda[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaZ.Series[0].AddXY(j, QuartisAlternativaZ[0,j]*100, formatfloat('0', QuartisAlternativaZ[0,j]*100) + '%');
          chrtalternativaZ.Series[1].AddXY(j, QuartisAlternativaZ[1,j]*100, formatfloat('0', QuartisAlternativaZ[1,j]*100) + '%');
          chrtalternativaZ.Series[2].AddXY(j, QuartisAlternativaZ[2,j]*100, formatfloat('0', QuartisAlternativaZ[2,j]*100) + '%');
          chrtalternativaZ.Series[2].AddXY(j, QuartisAlternativaZ[3,j]*100, formatfloat('0', QuartisAlternativaZ[3,j]*100) + '%');
          chrtalternativaZ.Series[4].AddXY(j, QuartisAlternativaZ[4,j]*100, formatfloat('0', QuartisAlternativaZ[4,j]*100) + '%');
          //chrtalternativaZ.Series[1].AddXY(j, rankingsegunda[j]*100, formatfloat('0', rankingsegunda[j]*100) + '%');
          //chrtalternativaZ.Series[2].AddXY(j, rankingterceira[j]*100, formatfloat('0', rankingterceira[j]*100) + '%');

        //end;
      end;

      chrtalternativaZ2.Series[0].Clear;
      chrtalternativaZ2.Series[1].Clear;
      chrtalternativaZ2.Series[2].Clear;
      chrtalternativaZ2.Series[3].Clear;
      chrtalternativaZ2.Series[4].Clear;

      chrtalternativaZ2.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaZ2.Series[1].Marks.FontSeriesColor := true;
      chrtalternativaZ2.Series[2].Marks.FontSeriesColor := true;
      chrtalternativaZ2.Series[3].Marks.FontSeriesColor := true;
      chrtalternativaZ2.Series[4].Marks.FontSeriesColor := true;

      chrtalternativaZ2.BottomAxis.Items.Clear;
      chrtalternativaZ2.BottomAxis.Items.Add(25,'0.25');
      chrtalternativaZ2.BottomAxis.Items.Add(50,'0.50');
      chrtalternativaZ2.BottomAxis.Items.Add(75,'0.75');
      chrtalternativaZ2.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaZ2.Series[0].AddXY(j, PercentilAlternativaZ[0,j]*100, formatfloat('0', PercentilAlternativaZ[0,j]*100) + '%');
          chrtalternativaZ2.Series[1].AddXY(j, PercentilAlternativaZ[1,j]*100, formatfloat('0', PercentilAlternativaZ[1,j]*100) + '%');
          chrtalternativaZ2.Series[2].AddXY(j, PercentilAlternativaZ[2,j]*100, formatfloat('0', PercentilAlternativaZ[2,j]*100) + '%');
          chrtalternativaZ2.Series[3].AddXY(j, PercentilAlternativaZ[3,j]*100, formatfloat('0', PercentilAlternativaZ[3,j]*100) + '%');
          chrtalternativaZ2.Series[4].AddXY(j, PercentilAlternativaZ[4,j]*100, formatfloat('0', PercentilAlternativaZ[4,j]*100) + '%');
          //chrtalternativaY.Series[1].AddXY(j, rankingsegunda[j]*100, formatfloat('0', rankingsegunda[j]*100) + '%');
          //chrtalternativaY.Series[2].AddXY(j, rankingterceira[j]*100, formatfloat('0', rankingterceira[j]*100) + '%');

        //end;
      end;

      imgalternativaZ.Picture.Bitmap.Assign
        (chrtalternativaZ2.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgalternativaZ.Width,
            imgalternativaZ.Height)));

      grdalternativaZ.Clear;
      grdalternativaZ.ColumnCount := 5;
      grdalternativaZ.RowCount := 6;

      for j := 0 to grdalternativaZ.ColumnCount - 1 do
        for I := 0 to grdalternativaZ.RowCount - 1 do
        begin
          grdalternativaZ.Cell[I, j].Alignment := taCenter;
          grdalternativaZ.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdalternativaZ.ColumnCount - 1 do
        grdalternativaZ.Cell[0, j].Font.Style := [fsbold];

      grdalternativaZ.Cell[0, 0].Text := 'Position';
      grdalternativaZ.Cell[1, 0].Text := '1º';
      grdalternativaZ.Cell[2, 0].Text := '2º';
      grdalternativaZ.Cell[3, 0].Text := '3º';
      grdalternativaZ.Cell[4, 0].Text := '1º or 2º';
      grdalternativaZ.Cell[5, 0].Text := '1º or 2º or 3º';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdalternativaZ.Cell[0, 1].Text := 'Percentile = 0.25';
      grdalternativaZ.Cell[0, 2].Text := 'Percentile = 0.50';
      grdalternativaZ.Cell[0, 3].Text := 'Percentile = 0.75';
      grdalternativaZ.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      for j := 0 to 4 do
      begin
        grdalternativaZ.Cell[j+1, i+1].Text := formatfloat('0', QuartisAlternativaZ[j,i] * 100) + '%';
        //grdalternativaZ.Cell[2, j+1].Text := formatfloat('0', MatrizFaixasegunda[1,j] * 100) + '%';
        //grdalternativaZ.Cell[3, j+1].Text := formatfloat('0', MatrizFaixasegunda[2,j] * 100) + '%';

      end;



      chrtalternativaY.Series[0].Clear;
      chrtalternativaY.Series[1].Clear;

      chrtalternativaY.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaY.Series[1].Marks.FontSeriesColor := true;


      chrtalternativaY.BottomAxis.Items.Clear;
      chrtalternativaY.BottomAxis.Items.Add(25,'0.25');
      chrtalternativaY.BottomAxis.Items.Add(50,'0.50');
      chrtalternativaY.BottomAxis.Items.Add(75,'0.75');
      chrtalternativaY.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaY.Series[0].AddXY(j, PercentilalternativaY[0,j]*100, formatfloat('0', PercentilalternativaY[0,j]*100) + '%');
          chrtalternativaY.Series[1].AddXY(j, PercentilalternativaY[1,j]*100, formatfloat('0', PercentilalternativaY[1,j]*100) + '%');

        //end;
      end;
      imgalternativaY.Picture.Bitmap.Assign
        (chrtalternativaY.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgalternativaY.Width,
            imgalternativaY.Height)));

      grdalternativaY.Clear;
      grdalternativaY.ColumnCount := 5;
      grdalternativaY.RowCount := 3;

      for j := 0 to grdalternativaY.ColumnCount - 1 do
        for I := 0 to grdalternativaY.RowCount - 1 do
        begin
          grdalternativaY.Cell[I, j].Alignment := taCenter;
          grdalternativaY.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdalternativaY.ColumnCount - 1 do
        grdalternativaY.Cell[0, j].Font.Style := [fsbold];

      grdalternativaY.Cell[0, 0].Text := 'Type of Probability';
      grdalternativaY.Cell[1, 0].Text := 'Probability of Z in 2º and X in 1º';
      grdalternativaY.Cell[2, 0].Text := 'Probability of Z in 2º';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdalternativaY.Cell[0, 1].Text := 'Percentile = 0.25';
      grdalternativaY.Cell[0, 2].Text := 'Percentile = 0.50';
      grdalternativaY.Cell[0, 3].Text := 'Percentile = 0.75';
      grdalternativaY.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      for j := 0 to 1 do
      begin
        grdalternativaY.Cell[j+1, i+1].Text := formatfloat('0', QuartisalternativaY[j,i] * 100) + '%';
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;

      chrtalternativaW.Series[0].Clear;
      chrtalternativaW.Series[1].Clear;

      chrtalternativaW.Series[0].Marks.FontSeriesColor := true;
      chrtalternativaW.Series[1].Marks.FontSeriesColor := true;


      chrtalternativaW.BottomAxis.Items.Clear;
      chrtalternativaW.BottomAxis.Items.Add(25,'0.25');
      chrtalternativaW.BottomAxis.Items.Add(50,'0.50');
      chrtalternativaW.BottomAxis.Items.Add(75,'0.75');
      chrtalternativaW.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtalternativaW.Series[0].AddXY(j, PercentilalternativaW[0,j]*100, formatfloat('0', PercentilalternativaW[0,j]*100) + '%');
          chrtalternativaW.Series[1].AddXY(j, PercentilalternativaW[1,j]*100, formatfloat('0', PercentilalternativaW[1,j]*100) + '%');

        //end;
      end;
      imgalternativaW.Picture.Bitmap.Assign
        (chrtalternativaW.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgalternativaW.Width,
            imgalternativaW.Height)));

      grdalternativaW.Clear;
      grdalternativaW.ColumnCount := 5;
      grdalternativaW.RowCount := 3;

      for j := 0 to grdalternativaW.ColumnCount - 1 do
        for I := 0 to grdalternativaW.RowCount - 1 do
        begin
          grdalternativaW.Cell[I, j].Alignment := taCenter;
          grdalternativaW.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdalternativaW.ColumnCount - 1 do
        grdalternativaW.Cell[0, j].Font.Style := [fsbold];

      grdalternativaW.Cell[0, 0].Text := 'Type of Probability';
      grdalternativaW.Cell[1, 0].Text := 'Probability of W in 2º and X in 1º';
      grdalternativaW.Cell[2, 0].Text := 'Probability of W in 2º';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdalternativaW.Cell[0, 1].Text := 'Percentile = 0.25';
      grdalternativaW.Cell[0, 2].Text := 'Percentile = 0.50';
      grdalternativaW.Cell[0, 3].Text := 'Percentile = 0.75';
      grdalternativaW.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      for j := 0 to 1 do
      begin
        grdalternativaW.Cell[j+1, i+1].Text := formatfloat('0', QuartisalternativaW[j,i] * 100) + '%';
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;


      // Resultados da Perda Máxima
      chrtperdamaxima.Series[0].Clear;
      chrtperdamaxima.Series[1].Clear;
      chrtperdamaxima.Series[2].Clear;

      chrtperdamaxima.Series[0].Marks.FontSeriesColor := true;
      chrtperdamaxima.Series[1].Marks.FontSeriesColor := true;
      chrtperdamaxima.Series[2].Marks.FontSeriesColor := true;

      chrtperdamaxima.BottomAxis.Items.Clear;
      chrtperdamaxima.BottomAxis.Items.Add(25,'0.25');
      chrtperdamaxima.BottomAxis.Items.Add(50,'0.50');
      chrtperdamaxima.BottomAxis.Items.Add(75,'0.75');
      chrtperdamaxima.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtperdamaxima.Series[0].AddXY(j, PercentilPerdaMaximaX[j], formatfloat('0.00', PercentilPerdaMaximaX[j]));
          chrtperdamaxima.Series[1].AddXY(j, PercentilPerdaMaximaZ[j], formatfloat('0.00', PercentilPerdaMaximaZ[j]));
          chrtperdamaxima.Series[2].AddXY(j, PercentilPerdaMaximaS[j], formatfloat('0.00', PercentilPerdaMaximaS[j]));
        //end;
      end;
      imgperdamaxima.Picture.Bitmap.Assign
        (chrtperdamaxima.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgperdamaxima.Width,
            imgperdamaxima.Height)));

      grdperdamaxima.Clear;
      grdperdamaxima.ColumnCount := 5;
      grdperdamaxima.RowCount := 4;

      for j := 0 to grdperdamaxima.ColumnCount - 1 do
        for I := 0 to grdperdamaxima.RowCount - 1 do
        begin
          grdperdamaxima.Cell[I, j].Alignment := taCenter;
          grdperdamaxima.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdperdamaxima.ColumnCount - 1 do
        grdperdamaxima.Cell[0, j].Font.Style := [fsbold];

      grdperdamaxima.Cell[0, 0].Text := 'Alternative of Interest';
      grdperdamaxima.Cell[1, 0].Text := 'Alternative X';
      grdperdamaxima.Cell[2, 0].Text := 'Alternative Z';
      grdperdamaxima.Cell[2, 0].Text := 'Alternative S';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdperdamaxima.Cell[0, 1].Text := 'Percentile = 0.25';
      grdperdamaxima.Cell[0, 2].Text := 'Percentile = 0.50';
      grdperdamaxima.Cell[0, 3].Text := 'Percentile = 0.75';
      grdperdamaxima.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      begin
        grdperdamaxima.Cell[1, i+1].Text := formatfloat('0.00', QuartisPerdaMaximaX[i]);
        grdperdamaxima.Cell[2, i+1].Text := formatfloat('0.00', QuartisPerdaMaximaZ[i]);
        grdperdamaxima.Cell[3, i+1].Text := formatfloat('0.00', QuartisPerdaMaximaS[i]);
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;

      // Resultados do Consequence Range
      chrtConsequenceRange.Series[0].Clear;
      chrtConsequenceRange.Series[1].Clear;
      //chrtConsequenceRange.Series[2].Clear;

      chrtConsequenceRange.Series[0].Marks.FontSeriesColor := true;
      chrtConsequenceRange.Series[1].Marks.FontSeriesColor := true;
      //chrtConsequenceRange.Series[2].Marks.FontSeriesColor := true;

      chrtConsequenceRange.BottomAxis.Items.Clear;
      chrtConsequenceRange.BottomAxis.Items.Add(25,'0.25');
      chrtConsequenceRange.BottomAxis.Items.Add(50,'0.50');
      chrtConsequenceRange.BottomAxis.Items.Add(75,'0.75');
      chrtConsequenceRange.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtConsequenceRange.Series[0].AddXY(j, PercentilConsequenceRangeX[j], formatfloat('0.00', PercentilConsequenceRangeX[j]));
          //chrtConsequenceRange.Series[1].AddXY(j, PercentilConsequenceRangeZ[j], formatfloat('0.00', PercentilConsequenceRangeZ[j]));
          chrtConsequenceRange.Series[1].AddXY(j, PercentilConsequenceRangeS[j], formatfloat('0.00', PercentilConsequenceRangeS[j]));
        //end;
      end;
      imgConsequenceRange.Picture.Bitmap.Assign
        (chrtConsequenceRange.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgConsequenceRange.Width,
            imgConsequenceRange.Height)));

      grdConsequenceRange.Clear;
      grdConsequenceRange.ColumnCount := 5;
      grdConsequenceRange.RowCount := 3;

      for j := 0 to grdConsequenceRange.ColumnCount - 1 do
        for I := 0 to grdConsequenceRange.RowCount - 1 do
        begin
          grdConsequenceRange.Cell[I, j].Alignment := taCenter;
          grdConsequenceRange.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdConsequenceRange.ColumnCount - 1 do
        grdConsequenceRange.Cell[0, j].Font.Style := [fsbold];

      grdConsequenceRange.Cell[0, 0].Text := 'Alternative of Interest';
      grdConsequenceRange.Cell[1, 0].Text := 'Alternative X';
     // grdConsequenceRange.Cell[2, 0].Text := 'Alternative Z';
      grdConsequenceRange.Cell[2, 0].Text := 'Alternative S';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grdConsequenceRange.Cell[0, 1].Text := 'Percentile = 0.25';
      grdConsequenceRange.Cell[0, 2].Text := 'Percentile = 0.50';
      grdConsequenceRange.Cell[0, 3].Text := 'Percentile = 0.75';
      grdConsequenceRange.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      begin
        grdConsequenceRange.Cell[1, i+1].Text := formatfloat('0.00', QuartisConsequenceRangeX[i]);
        //grdConsequenceRange.Cell[2, i+1].Text := formatfloat('0.00', QuartisConsequenceRangeZ[i]);
        grdConsequenceRange.Cell[2, i+1].Text := formatfloat('0.00', QuartisConsequenceRangeS[i]);
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;


      // Resultados do Consequence Range
      chrtdifvg.Series[0].Clear;
      chrtdifvg.Series[1].Clear;
      //chrtdifvg.Series[2].Clear;

      chrtdifvg.Series[0].Marks.FontSeriesColor := true;
      chrtdifvg.Series[1].Marks.FontSeriesColor := true;
      //chrtdifvg.Series[2].Marks.FontSeriesColor := true;

      chrtdifvg.BottomAxis.Items.Clear;
      chrtdifvg.BottomAxis.Items.Add(25,'0.25');
      chrtdifvg.BottomAxis.Items.Add(50,'0.50');
      chrtdifvg.BottomAxis.Items.Add(75,'0.75');
      chrtdifvg.BottomAxis.Items.Add(100,'1.00');
      for j := 0 to 100 do
      begin
        //if (rankingprimeira[j] <> 0) and (round(rankingprimeira[j]*100) >= 1) {or (rankingsegunda[j] <> 0) or (rankingterceira[j] <> 0)} then
        //begin
          chrtdifvg.Series[0].AddXY(j, percentildifVGX[j], formatfloat('0.00', percentildifVGX[j]));
          //chrtConsequenceRange.Series[1].AddXY(j, PercentilConsequenceRangeZ[j], formatfloat('0.00', PercentilConsequenceRangeZ[j]));
          chrtdifvg.Series[1].AddXY(j, percentildifVGS[j], formatfloat('0.00', percentildifVGS[j]));
        //end;
      end;
      imgdifvg.Picture.Bitmap.Assign
        (chrtdifvg.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgdifvg.Width,
            imgdifvg.Height)));

      grddifvg.Clear;
      grddifvg.ColumnCount := 5;
      grddifvg.RowCount := 3;

      for j := 0 to grddifvg.ColumnCount - 1 do
        for I := 0 to grddifvg.RowCount - 1 do
        begin
          grddifvg.Cell[I, j].Alignment := taCenter;
          grddifvg.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grddifvg.ColumnCount - 1 do
        grddifvg.Cell[0, j].Font.Style := [fsbold];

      grddifvg.Cell[0, 0].Text := 'Alternative of Interest';
      grddifvg.Cell[1, 0].Text := 'Alternative X';
     // grddifvg.Cell[2, 0].Text := 'Alternative Z';
      grddifvg.Cell[2, 0].Text := 'Alternative S';

      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      //grdranking.Cell[0, 1].Text := '0% to 40%';
      grddifvg.Cell[0, 1].Text := 'Percentile = 0.25';
      grddifvg.Cell[0, 2].Text := 'Percentile = 0.50';
      grddifvg.Cell[0, 3].Text := 'Percentile = 0.75';
      grddifvg.Cell[0, 4].Text := 'Percentile = 1.00';
     // grdranking.Cell[0, 8].Text := 'Total';

      for I := 0 to 3 do
      begin
        grddifvg.Cell[1, i+1].Text := formatfloat('0.00', quartisdifVGX[i]);
        //grddifvg.Cell[2, i+1].Text := formatfloat('0.00', QuartisConsequenceRangeZ[i]);
        grddifvg.Cell[2, i+1].Text := formatfloat('0.00', quartisdifVGS[i]);
        //grdalternativaX.Cell[2, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[1,j] * 100) + '%';
        //grdalternativaX.Cell[3, j+1].Text := formatfloat('0', MatrizFaixaPrimeira[2,j] * 100) + '%';

      end;

      cmbbxmatrizconseqAltXS.ItemIndex := 0;
      cmbbxmatrizconseqAltXSChange(self);

      


      {
      grdregra.Clear;
      grdregra.ColumnCount := 6;
      grdregra.RowCount := 2;

      for j := 0 to grdregra.ColumnCount - 1 do
        for I := 0 to grdregra.RowCount - 1 do
        begin
          grdregra.Cell[I, j].Alignment := taCenter;
          grdregra.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdregra.ColumnCount - 1 do
        grdregra.Cell[0, j].Font.Style := [fsbold];

      grdregra.Cell[0, 0].Text := 'Rule';
      grdregra.Cell[0, 1].Text := 'Mean';
      grdregra.Cell[0, 2].Text := 'Standart Deviation';
      grdregra.Cell[0, 3].Text := 'Maximum';
      grdregra.Cell[0, 4].Text := 'Median';
      grdregra.Cell[0, 5].Text := 'Mode';

      grdregra.Cell[1, 0].Text := '1ª Alternative';
      grdregra.Cell[1, 1].Text := formatfloat('0.00', mediadesempenho);
      grdregra.Cell[1, 2].Text := formatfloat('0.00', DesvioPadrao);
      grdregra.Cell[1, 3].Text := formatfloat('0.00', maximodesempenho);
      grdregra.Cell[1, 4].Text := formatfloat('0.00', mediana);
      grdregra.Cell[1, 5].Text := formatfloat('0.00', moda);   }

      grdresultadosprobabilidade.Clear;
      grdresultadosprobabilidade.ColumnCount := 7;
      grdresultadosprobabilidade.RowCount := 4;

      for j := 0 to grdresultadosprobabilidade.ColumnCount - 1 do
        for I := 0 to grdresultadosprobabilidade.RowCount - 1 do
        begin
          grdresultadosprobabilidade.Cell[I, j].Alignment := taCenter;
          grdresultadosprobabilidade.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdresultadosprobabilidade.ColumnCount - 1 do
        grdresultadosprobabilidade.Cell[0, j].Font.Style := [fsbold];

      grdresultadosprobabilidade.Cell[0, 0].Text := 'Rule';
      // grdresultadosprobabilidade.Cell[0, 1].Text := 'Mean';
      grdresultadosprobabilidade.Cell[0, 1].Text := '40% to 100%';
      grdresultadosprobabilidade.Cell[0, 2].Text := '50% to 100%';
      grdresultadosprobabilidade.Cell[0, 3].Text := '60% to 100%';
      grdresultadosprobabilidade.Cell[0, 4].Text := '70% to 100%';
      grdresultadosprobabilidade.Cell[0, 5].Text := '80% to 100%';
      grdresultadosprobabilidade.Cell[0, 6].Text := '90% to 100%';

      // grdresultadosprobabilidade.Cell[1, 0].Text := 'Frequency of 1ª Alternative';
      grdresultadosprobabilidade.Cell[1, 0].Text :=
        'Cumulative Frequency of Best Alternative';
      grdresultadosprobabilidade.Cell[2, 0].Text :=
        'Cumulative Frequency of Two Bests Alternatives';
      grdresultadosprobabilidade.Cell[3, 0].Text :=
        'Cumulative Frequency of Three Bests Alternatives';

      { grdresultadosprobabilidade.Cell[1, 0].Text := 'Frequency of Best Alternative';
        grdresultadosprobabilidade.Cell[2, 0].Text := 'Frequency of Two Bests Alternatives';
        grdresultadosprobabilidade.Cell[3, 0].Text := 'Frequency of Three Bests Alternatives';
        }

      grdmediaprobalt.Clear;
      grdmediaprobalt.ColumnCount := 3;
      grdmediaprobalt.RowCount := 4;

      for j := 0 to grdmediaprobalt.ColumnCount - 1 do
        for I := 0 to grdmediaprobalt.RowCount - 1 do
        begin
          grdmediaprobalt.Cell[I, j].Alignment := taCenter;
          grdmediaprobalt.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdmediaprobalt.ColumnCount - 1 do
        grdmediaprobalt.Cell[0, j].Font.Style := [fsbold];

      grdmediaprobalt.Cell[0, 0].Text := 'Rule';
      // grdmediaprobalt.Cell[0, 1].Text := 'Mean';
      grdmediaprobalt.Cell[0, 1].Text := 'Mean';
      grdmediaprobalt.Cell[0, 2].Text := 'Standart Deviation';

      grdmediaprobalt.Cell[1, 0].Text := 'Best Alternative';
      grdmediaprobalt.Cell[2, 0].Text := 'Two Bests Alternatives';
      grdmediaprobalt.Cell[3, 0].Text := 'Three Bests Alternatives';

      grdmediaprobalt.Cell[1, 1].Text := formatfloat('0', mediaprobalt) + '%';
      grdmediaprobalt.Cell[1, 2].Text := formatfloat
        ('0.00', DesvioPadraoprobalt);

      grdmediaprobalt.Cell[2, 1].Text := formatfloat('0', mediaprobalt2) + '%';
      grdmediaprobalt.Cell[2, 2].Text := formatfloat
        ('0.00', DesvioPadraoprobalt2);

      grdmediaprobalt.Cell[3, 1].Text := formatfloat('0', mediaprobalt3) + '%';
      grdmediaprobalt.Cell[3, 2].Text := formatfloat
        ('0.00', DesvioPadraoprobalt3);

      {
        grdresultadosprobabilidade.Cell[1, 1].Text := formatfloat
        ('0', faixa40probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 2].Text := formatfloat
        ('0', faixa50probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 3].Text := formatfloat
        ('0', faixa60probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 4].Text := formatfloat
        ('0', faixa70probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 5].Text := formatfloat
        ('0', faixa80probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 6].Text := formatfloat
        ('0', faixa90probalt) + '%'; }

      { grdresultadosprobabilidade.Cell[1, 1].Text := formatfloat
        ('0', faixa40probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 2].Text := formatfloat
        ('0', faixa50probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 3].Text := formatfloat
        ('0', faixa60probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 4].Text := formatfloat
        ('0', faixa70probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 5].Text := formatfloat
        ('0', faixa80probalt) + '%';
        grdresultadosprobabilidade.Cell[1, 6].Text := formatfloat
        ('0', faixa90probalt) + '%';
        }
      grdresultadosprobabilidade.Cell[1, 1].Text := formatfloat
        ('0', faixa40probaltacumulado) + '%';
      grdresultadosprobabilidade.Cell[1, 2].Text := formatfloat
        ('0', faixa50probaltacumulado) + '%';
      grdresultadosprobabilidade.Cell[1, 3].Text := formatfloat
        ('0', faixa60probaltacumulado) + '%';
      grdresultadosprobabilidade.Cell[1, 4].Text := formatfloat
        ('0', faixa70probaltacumulado) + '%';
      grdresultadosprobabilidade.Cell[1, 5].Text := formatfloat
        ('0', faixa80probaltacumulado) + '%';
      grdresultadosprobabilidade.Cell[1, 6].Text := formatfloat
        ('0', faixa90probaltacumulado) + '%';

      { grdresultadosprobabilidade.Cell[2, 1].Text := formatfloat
        ('0', faixa40probalt2) + '%';
        grdresultadosprobabilidade.Cell[2, 2].Text := formatfloat
        ('0', faixa50probalt2) + '%';
        grdresultadosprobabilidade.Cell[2, 3].Text := formatfloat
        ('0', faixa60probalt2) + '%';
        grdresultadosprobabilidade.Cell[2, 4].Text := formatfloat
        ('0', faixa70probalt2) + '%';
        grdresultadosprobabilidade.Cell[2, 5].Text := formatfloat
        ('0', faixa80probalt2) + '%';
        grdresultadosprobabilidade.Cell[2, 6].Text := formatfloat
        ('0', faixa90probalt2) + '%';
        }
      grdresultadosprobabilidade.Cell[2, 1].Text := formatfloat
        ('0', faixa40probaltacumulado2) + '%';
      grdresultadosprobabilidade.Cell[2, 2].Text := formatfloat
        ('0', faixa50probaltacumulado2) + '%';
      grdresultadosprobabilidade.Cell[2, 3].Text := formatfloat
        ('0', faixa60probaltacumulado2) + '%';
      grdresultadosprobabilidade.Cell[2, 4].Text := formatfloat
        ('0', faixa70probaltacumulado2) + '%';
      grdresultadosprobabilidade.Cell[2, 5].Text := formatfloat
        ('0', faixa80probaltacumulado2) + '%';
      grdresultadosprobabilidade.Cell[2, 6].Text := formatfloat
        ('0', faixa90probaltacumulado2) + '%';

      {
        grdresultadosprobabilidade.Cell[3, 1].Text := formatfloat
        ('0', faixa40probalt3) + '%';
        grdresultadosprobabilidade.Cell[3, 2].Text := formatfloat
        ('0', faixa50probalt3) + '%';
        grdresultadosprobabilidade.Cell[3, 3].Text := formatfloat
        ('0', faixa60probalt3) + '%';
        grdresultadosprobabilidade.cell[3, 4].Text := formatfloat
        ('0', faixa70probalt3) + '%';
        grdresultadosprobabilidade.Cell[3, 5].Text := formatfloat
        ('0', faixa80probalt3) + '%';
        grdresultadosprobabilidade.Cell[3, 6].Text := formatfloat
        ('0', faixa90probalt3) + '%';
        }
      grdresultadosprobabilidade.Cell[3, 1].Text := formatfloat
        ('0', faixa40probaltacumulado3) + '%';
      grdresultadosprobabilidade.Cell[3, 2].Text := formatfloat
        ('0', faixa50probaltacumulado3) + '%';
      grdresultadosprobabilidade.Cell[3, 3].Text := formatfloat
        ('0', faixa60probaltacumulado3) + '%';
      grdresultadosprobabilidade.Cell[3, 4].Text := formatfloat
        ('0', faixa70probaltacumulado3) + '%';
      grdresultadosprobabilidade.Cell[3, 5].Text := formatfloat
        ('0', faixa80probaltacumulado3) + '%';
      grdresultadosprobabilidade.Cell[3, 6].Text := formatfloat
        ('0', faixa90probaltacumulado3) + '%';

      TiwUserSession(WebApplication.Data).percentualr1 := regra1 / Ninst;
      TiwUserSession(WebApplication.Data).percentualr2 := regra2 / Ninst;
      TiwUserSession(WebApplication.Data).percentualr3 := regra3 / Ninst;
      TiwUserSession(WebApplication.Data).percentualr4 := regra4 / Ninst;

      { TiwUserSession(WebApplication.Data).mediaperdamedia := mediaperdamedia / ninst;
        TiwUserSession(WebApplication.Data).mediaperdamaxima := mediaperdamaxima / ninst;
        TiwUserSession(WebApplication.Data).maximaperdamedia := RoundTo(maximaperdamedia, -2);
        TiwUserSession(WebApplication.Data).maximaperdamaxima := roundto(maximaperdamaxima, -2);
        }
     { for I := 0 to 3 do
      begin
        for j := 0 to 4 do
        begin
          if probregra[I, j] <> 0 then
          begin
            TiwUserSession(WebApplication.Data).mediaperdamedia[I, j] :=
              mediaperdamedia[I, j] / TiwUserSession(WebApplication.Data)
              .probregra[I, j];
            TiwUserSession(WebApplication.Data).mediaperdamaxima[I, j] :=
              mediaperdamaxima[I, j] / probregra[I, j];
          end;
          if (I <> 3) and (j <> 4) then
          begin
            if maximaperdamedia[I, j] > maximaperdamedia[3, 4] then
              maximaperdamedia[3, 4] := maximaperdamedia[I, j];
            if maximaperdamaxima[I, j] > maximaperdamaxima[3, 4] then
              maximaperdamaxima[3, 4] := maximaperdamaxima[I, j];

          end;
        end;
      end;
      for I := 0 to 2 do
      begin
        if maximaperdamedia[I, 0] > maximaperdamedia[3, 0] then
          maximaperdamedia[3, 0] := maximaperdamedia[I, 0];
        if maximaperdamedia[I, 1] > maximaperdamedia[3, 1] then
          maximaperdamedia[3, 1] := maximaperdamedia[I, 1];
        if maximaperdamedia[I, 2] > maximaperdamedia[3, 2] then
          maximaperdamedia[3, 2] := maximaperdamedia[I, 2];
        if maximaperdamedia[I, 3] > maximaperdamedia[3, 3] then
          maximaperdamedia[3, 3] := maximaperdamedia[I, 3];

        if maximaperdamaxima[I, 0] > maximaperdamaxima[3, 0] then
          maximaperdamaxima[3, 0] := maximaperdamaxima[I, 0];
        if maximaperdamaxima[I, 1] > maximaperdamaxima[3, 1] then
          maximaperdamaxima[3, 1] := maximaperdamaxima[I, 1];
        if maximaperdamaxima[I, 2] > maximaperdamaxima[3, 2] then
          maximaperdamaxima[3, 2] := maximaperdamaxima[I, 2];
        if maximaperdamaxima[I, 3] > maximaperdamaxima[3, 3] then
          maximaperdamaxima[3, 3] := maximaperdamaxima[I, 3];
      end;

      for j := 0 to 3 do
      begin
        if maximaperdamedia[0, j] > maximaperdamedia[0, 4] then
          maximaperdamedia[0, 4] := maximaperdamedia[0, j];
        if maximaperdamedia[1, j] > maximaperdamedia[1, 4] then
          maximaperdamedia[1, 4] := maximaperdamedia[1, j];
        if maximaperdamedia[2, j] > maximaperdamedia[2, 4] then
          TiwUserSession(WebApplication.Data).maximaperdamedia[2, 4] :=
            maximaperdamedia[2, j];

        if maximaperdamaxima[0, j] > maximaperdamaxima[0, 4] then
          maximaperdamaxima[0, 4] := maximaperdamaxima[0, j];
        if maximaperdamaxima[1, j] > maximaperdamaxima[1, 4] then
          maximaperdamaxima[1, 4] := maximaperdamaxima[1, j];
        if maximaperdamaxima[2, j] > maximaperdamaxima[2, 4] then
          TiwUserSession(WebApplication.Data).maximaperdamaxima[2, 4] :=
            maximaperdamaxima[2, j];
      end;

      grdestatistica.Clear;
      grdestatistica.RowCount := 6;
      grdestatistica.ColumnCount := 6;

      for j := 0 to grdestatistica.ColumnCount - 1 do
        for I := 0 to grdestatistica.RowCount - 1 do
        begin
          grdestatistica.Cell[I, j].Alignment := taCenter;
          grdestatistica.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdestatistica.ColumnCount - 1 do
        grdestatistica.Cell[0, j].Font.Style := [fsbold];

      for I := 0 to grdestatistica.RowCount - 1 do
        grdestatistica.Cell[I, 0].Font.Style := [fsbold];

      grdestatistica.Cell[2, 0].Text := '1 Alternative';
      grdestatistica.Cell[3, 0].Text := '2 Alternatives';
      grdestatistica.Cell[4, 0].Text := '3 Alternatives';
      grdestatistica.Cell[5, 0].Text := 'Overall';

      grdestatistica.Cell[0, 1].Text := '> 80%';
      grdestatistica.Cell[0, 2].Text := '70% to 80%';
      grdestatistica.Cell[0, 3].Text := '60% to 70%';
      grdestatistica.Cell[0, 4].Text := '50% to 60%';
      grdestatistica.Cell[0, 5].Text := 'Overall';

      for j := 1 to 5 do
        grdestatistica.Cell[1, j].Text := 'μ | max';

      for I := 2 to grdestatistica.RowCount - 1 do
        for j := 1 to grdestatistica.ColumnCount - 1 do
        begin
          grdestatistica.Cell[I, j].Text := formatfloat
            ('0.00', mediaperdamedia[I - 2, j - 1]) + ' | ' + formatfloat
            ('0.00', maximaperdamedia[I - 2, j - 1]);
        end;

      grdmaxmin.Clear;
      grdmaxmin.RowCount := 6;
      grdmaxmin.ColumnCount := 6;

      for j := 0 to grdmaxmin.ColumnCount - 1 do
        for I := 0 to grdmaxmin.RowCount - 1 do
        begin
          grdmaxmin.Cell[I, j].Alignment := taCenter;
          grdmaxmin.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdmaxmin.ColumnCount - 1 do
        grdmaxmin.Cell[0, j].Font.Style := [fsbold];

      for I := 0 to grdmaxmin.RowCount - 1 do
        grdmaxmin.Cell[I, 0].Font.Style := [fsbold];

      grdmaxmin.Cell[2, 0].Text := '1 Alternative';
      grdmaxmin.Cell[3, 0].Text := '2 Alternatives';
      grdmaxmin.Cell[4, 0].Text := '3 Alternatives';
      grdmaxmin.Cell[5, 0].Text := 'Overall';

      grdmaxmin.Cell[0, 1].Text := '> 80%';
      grdmaxmin.Cell[0, 2].Text := '70% to 80%';
      grdmaxmin.Cell[0, 3].Text := '60% to 70%';
      grdmaxmin.Cell[0, 4].Text := '50% to 60%';
      grdmaxmin.Cell[0, 5].Text := 'Overall';

      for j := 1 to 5 do
        grdmaxmin.Cell[1, j].Text := 'μ | max';

      for I := 2 to grdmaxmin.RowCount - 1 do
        for j := 1 to grdmaxmin.ColumnCount - 1 do
        begin
          grdmaxmin.Cell[I, j].Text := formatfloat
            ('0.00', mediaperdamaxima[I - 2, j - 1]) + ' | ' + formatfloat
            ('0.00', maximaperdamaxima[I - 2, j - 1]);
        end;   }

      grdregradecisao.ColumnCount := 2;
      grdregradecisao.RowCount := 5;
      // grdsegundo.ColumnCount := 1 + Numcrit;

      for j := 0 to grdregradecisao.ColumnCount - 1 do
        for I := 0 to grdregradecisao.RowCount - 1 do
        begin
          grdregradecisao.Cell[I, j].Alignment := taCenter;
          grdregradecisao.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdregradecisao.ColumnCount - 1 do
        grdregradecisao.Cell[0, j].Font.Style := [fsbold];

      grdregradecisao.Cell[0, 0].Text := 'Recomendation';
      grdregradecisao.Cell[1, 0].Text := '1 Alternative';
      grdregradecisao.Cell[2, 0].Text := '2 Alternatives';
      grdregradecisao.Cell[3, 0].Text := '3 Alternatives';
      grdregradecisao.Cell[4, 0].Text := 'Not possible';

      grdregradecisao.Cell[0, 1].Text := 'Probability';
      grdregradecisao.Cell[1, 1].Text := formatfloat('0', percentualr1 * 100)
        + '%';
      grdregradecisao.Cell[2, 1].Text := formatfloat('0', percentualr2 * 100)
        + '%';
      grdregradecisao.Cell[3, 1].Text := formatfloat('0', percentualr3 * 100)
        + '%';
      grdregradecisao.Cell[4, 1].Text := formatfloat('0', percentualr4 * 100)
        + '%';

      QtdMedia := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).QtdMedia := QtdMedia + (I + 1)
            * NsolInst[I] / Ninst;
          inc(j);
        end;
      lblmedia.Caption := formatfloat('0.00', QtdMedia);

      mediana := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          j := j + round(NsolInst[I]);
          if j = round(Ninst / 2) then
          begin
            mediana := ((I + 1) + (I + 2)) / 2;
            Break;
          end;
          if j > round(Ninst / 2) then
          begin
            mediana := I + 1;
            Break;
          end;
        end;
      lblmediana.Caption := floattostr(mediana);

      moda := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          if j < NsolInst[I] then
          begin
            j := round(NsolInst[I]);
            moda := I + 1;
          end;

        end;

      lblmoda.Caption := floattostr(moda);

      DesvioPadrao := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadrao :=
            DesvioPadrao + NsolInst[I] * Sqr(abs((I + 1) - QtdMedia))
        end;
      DesvioPadrao := Sqrt(DesvioPadrao / Ninst);

      lbldesviopadrao.Caption := formatfloat('0.00', DesvioPadrao);

    end;
    if sobreclassificacao then
    begin
      TiwUserSession(WebApplication.Data).percentualr1Promethee :=
        regra1Promethee / Ninst;
      TiwUserSession(WebApplication.Data).percentualr2Promethee :=
        regra2Promethee / Ninst;
      TiwUserSession(WebApplication.Data).percentualr3Promethee :=
        regra3Promethee / Ninst;
      TiwUserSession(WebApplication.Data).percentualr4Promethee :=
        regra4Promethee / Ninst;

      grdregradecisaopromethee.ColumnCount := 2;
      grdregradecisaopromethee.RowCount := 5;
      // grdsegundo.ColumnCount := 1 + Numcrit;

      for j := 0 to grdregradecisaopromethee.ColumnCount - 1 do
        for I := 0 to grdregradecisaopromethee.RowCount - 1 do
        begin
          grdregradecisaopromethee.Cell[I, j].Alignment := taCenter;
          grdregradecisaopromethee.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdregradecisaopromethee.ColumnCount - 1 do
        grdregradecisaopromethee.Cell[0, j].Font.Style := [fsbold];

      grdregradecisaopromethee.Cell[0, 0].Text := 'Recomendation';
      grdregradecisaopromethee.Cell[1, 0].Text := '1 Alternative';
      grdregradecisaopromethee.Cell[2, 0].Text := '2 Alternatives';
      grdregradecisaopromethee.Cell[3, 0].Text := '3 Alternatives';
      grdregradecisaopromethee.Cell[4, 0].Text := 'Not possible';

      grdregradecisaopromethee.Cell[0, 1].Text := 'Probability';
      grdregradecisaopromethee.Cell[1, 1].Text := formatfloat
        ('0', percentualr1Promethee * 100) + '%';
      grdregradecisaopromethee.Cell[2, 1].Text := formatfloat
        ('0', percentualr2Promethee * 100) + '%';
      grdregradecisaopromethee.Cell[3, 1].Text := formatfloat
        ('0', percentualr3Promethee * 100) + '%';
      grdregradecisaopromethee.Cell[4, 1].Text := formatfloat
        ('0', percentualr4Promethee * 100) + '%';

      QtdMediaPromethee := 0;
      j := 0;
      for I := 0 to Length(NsolInstPromethee) - 1 do
        if NsolInstPromethee[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).QtdMediaPromethee :=
            QtdMediaPromethee + (I + 1) * NsolInstPromethee[I] / Ninst;
          inc(j);
        end;
      lblmediapromethee.Caption := formatfloat('0.00', QtdMediaPromethee);

      medianaPromethee := 0;
      j := 0;
      for I := 0 to Length(NsolInstPromethee) - 1 do
        if NsolInstPromethee[I] <> 0 then
        begin
          j := j + round(NsolInstPromethee[I]);
          if j = round(Ninst / 2) then
          begin
            medianaPromethee := ((I + 1) + (I + 2)) / 2;
            Break;
          end;
          if j > round(Ninst / 2) then
          begin
            medianaPromethee := I + 1;
            Break;
          end;
        end;
      lblmedianapromethee.Caption := floattostr(medianaPromethee);

      modaPromethee := 0;
      j := 0;
      for I := 0 to Length(NsolInstPromethee) - 1 do
        if NsolInstPromethee[I] <> 0 then
        begin
          if j < NsolInstPromethee[I] then
          begin
            j := round(NsolInstPromethee[I]);
            modaPromethee := I + 1;
          end;

        end;

      lblmodapromethee.Caption := floattostr(modaPromethee);

      DesvioPadraoPromethee := 0;
      for I := 0 to Length(NsolInstPromethee) - 1 do
        if NsolInstPromethee[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadraoPromethee :=
            DesvioPadraoPromethee + NsolInstPromethee[I] * Sqr
            (abs((I + 1) - QtdMediaPromethee))
        end;
      DesvioPadraoPromethee := Sqrt(DesvioPadraoPromethee / Ninst);

      lbldesviopadraopromethee.Caption := formatfloat
        ('0.00', DesvioPadraoPromethee);
    end;
    {
      TiwUserSession(WebApplication.Data).TotalPrimCrit := 0;
      for I := 0 to numcrit - 1 do
      TotalPrimCrit := TotalPrimCrit + TiwUserSession(WebApplication.Data)
      .Primcrit[I];

      TotalPrimCrit := (TotalPrimCrit) / (numcrit * Ninst);

      if numcrit > 3 then
      begin
      TiwUserSession(WebApplication.Data).totalseccrit := 0;
      for I := 0 to numcrit - 1 do
      totalseccrit := totalseccrit + TiwUserSession(WebApplication.Data)
      .Seccrit[I];

      totalseccrit := (totalseccrit) / (numcrit * (numcrit - 1) * Ninst);

      end;

      if numcrit > 4 then
      begin
      TiwUserSession(WebApplication.Data).totalterccrit := 0;
      for I := 0 to numcrit - 1 do
      totalterccrit := totalterccrit + TiwUserSession(WebApplication.Data)
      .terccrit[I];

      totalterccrit := (totalterccrit) /
      (numcrit * (numcrit - 1) * (numcrit - 2) * Ninst);
      end;

      if numcrit > 6 then
      begin
      TiwUserSession(WebApplication.Data).totalquartcrit := 0;
      for I := 0 to numcrit - 1 do
      totalquartcrit := totalquartcrit + TiwUserSession(WebApplication.Data)
      .quartcrit[I];
      totalquartcrit := (totalquartcrit) /
      (numcrit * (numcrit - 1) * (numcrit - 2) * (numcrit - 3) * Ninst);
      end;
      }
    SetarCores;
    if aditivo then
    begin
      chrtResultado.RemoveAllSeries;
      // chrtresultado.ClearChart;
      With TiwUserSession(WebApplication.Data) do
      begin
        j := 0;
        SetLength(bar, j + 1);
        bar[j] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[j]);
        for I := 0 to Length(NsolInst) - 1 do
        begin
          if round(TiwUserSession(WebApplication.Data).NsolInst[I]) <> 0 then
          begin
            bar[0].AddXY
              (j * 2, (NsolInst[I] / TiwUserSession(WebApplication.Data).Ninst)
                * 100, formatfloat('0.00', (NsolInst[I] / Ninst) * 100) + '%');
            chrtResultado.BottomAxis.Items.Add(j * 2, inttostr(I + 1));
            // chrtResultado.Legend.Item[0].Text := inttostr(i + 1) + ' Solutions ';
            inc(j);
          end;
        end;
      end;

      chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.BottomAxis.visible := false;
      chrtResultado.Legend.Visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Maximum := 100;
      chrtResultado.LeftAxis.MaximumOffset := 16;
      chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.LeftAxis.AutomaticMaximum := false;
      chrtResultado.LeftAxis.Increment := 20;
      chrtResultado.Title.Caption :=
        'Percentage of cases with X number of solutions';
      imgResultado.Picture.Bitmap.Assign(chrtResultado.TeeCreateBitmap
          (clWebPURPLE, rect(0, 0, imgResultado.Width, imgResultado.Height)));
      if numcrit <= 3 then
        grdresultados.RowCount := 3;
      if numcrit > 3 then
        grdresultados.RowCount := 4;
      if numcrit > 4 then
        grdresultados.RowCount := 5;
      if numcrit > 6 then
        grdresultados.RowCount := 6;




      // grdsegundo.RowCount := 6;

      grdresultados.ColumnCount := 2;
      // grdsegundo.ColumnCount := 1 + Numcrit;

      for j := 0 to grdresultados.ColumnCount - 1 do
        for I := 0 to grdresultados.RowCount - 1 do
        begin
          grdresultados.Cell[I, j].Alignment := taCenter;
          grdresultados.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      for j := 0 to grdresultados.ColumnCount - 1 do
        grdresultados.Cell[0, j].Font.Style := [fsbold];

      for I := 0 to grdresultados.RowCount - 1 do
        grdresultados.Cell[I, 0].Font.Style := [fsbold];

      Fatorial := 1;
      for I := 1 to numcrit - 1 do
        Fatorial := Fatorial * I;
      // for i := 0 to Numcrit - 1 do
      // begin
      I := 0;
      grdresultados.Cell[0, I + 1].Text := 'Percentage';
      grdresultados.Cell[1, I + 1].Text := formatfloat
        ('0.00', TotalPrimCrit * 100) + ' %';
      if numcrit > 3 then
      begin
        grdresultados.Cell[2, I + 1].Text := formatfloat
          ('0.00', totalseccrit * 100) + ' %';
        grdresultados.Cell[3, I + 1].Text := formatfloat('0.00', 1 * 100)
          + ' %';
      end
      else
        grdresultados.Cell[2, I + 1].Text := formatfloat('0.00', 1 * 100)
          + ' %';
      if numcrit > 4 then
      begin
        grdresultados.Cell[3, I + 1].Text := formatfloat
          ('0.00', totalterccrit * 100) + ' %';
        grdresultados.Cell[4, I + 1].Text := formatfloat('0.00', 1 * 100)
          + ' %';
      end;
      if numcrit = 4 then
        grdresultados.Cell[3, I + 1].Text := formatfloat('0.00', 1 * 100)
          + ' %';
      if numcrit > 6 then
      begin
        grdresultados.Cell[4, I + 1].Text := formatfloat
          ('0.00', totalquartcrit * 100) + ' %';
        grdresultados.Cell[5, I + 1].Text := formatfloat
          ('0.00', Fatorial / Fatorial * 100) + ' %';
      end;

      // end;

      grdresultados.Cell[0, 0].Text := 'Rank Position:';
      grdresultados.Cell[1, 0].Text := '1th:';
      grdresultados.Cell[2, 0].Text := 'Others:';
      if numcrit = 4 then
      begin
        grdresultados.Cell[2, 0].Text := '1th + 2th:';
        grdresultados.Cell[3, 0].Text := 'Others:';
      end;
      if (numcrit = 5) or (numcrit = 6) then
      begin
        grdresultados.Cell[2, 0].Text := '1th + 2th:';
        grdresultados.Cell[3, 0].Text := '1th + 2th + 3th:';
        grdresultados.Cell[4, 0].Text := 'Others:';
      end;
      if numcrit > 6 then
      begin
        grdresultados.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
        grdresultados.Cell[5, 0].Text := 'Others:';
      end;

      grdnsolutions.RowCount := 2;
      grdnsolutions.ColumnCount := 2;

      for j := 0 to grdnsolutions.ColumnCount - 1 do
        grdnsolutions.Cell[0, j].Font.Style := [fsbold];

      // for i := 0 to grdresultados.RowCount - 1 do
      // grdnsolutions.Cell[i,0].Font.Style := [fsbold];

      grdnsolutions.Cell[0, 0].Text := 'Number of Solutions';
      grdnsolutions.Cell[0, 1].Text := 'Percentage';
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
      begin
        if round(NsolInst[I]) = 0 then
          continue
        else
        begin
          grdnsolutions.RowCount := 2 + j;
          if I = 0 then
            grdnsolutions.Cell[1 + j, 0].Text := '1 Solution'
          else
            grdnsolutions.Cell[j + 1, 0].Text := inttostr(I + 1) + ' Solutions';
          grdnsolutions.Cell[j + 1, 1].Text := formatfloat
            ('0.00', (NsolInst[I] / Ninst) * 100) + '%';
          inc(j);
        end;
      end;
      for j := 0 to grdnsolutions.ColumnCount - 1 do
        for I := 0 to grdnsolutions.RowCount - 1 do
        begin
          grdnsolutions.Cell[I, j].Alignment := taCenter;
          grdnsolutions.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      rgnmatcons.Visible := false;
      rgnresultado.Visible := true;
      // btnshowresults.Enabled := true;

    end;
    if sobreclassificacao then
    begin
      chrtresultadopromethee.RemoveAllSeries;
      // chrtresultado.ClearChart;
      With TiwUserSession(WebApplication.Data) do
      begin
        j := 0;
        SetLength(barpromethee, j + 1);
        barpromethee[j] := TBarSeries.Create(self);
        chrtresultadopromethee.AddSeries(barpromethee[j]);
        for I := 0 to Length(NsolInstPromethee) - 1 do
        begin
          if round(TiwUserSession(WebApplication.Data).NsolInstPromethee[I])
            <> 0 then
          begin
            barpromethee[0].AddXY(j * 2, (NsolInstPromethee[I] / TiwUserSession
                  (WebApplication.Data).Ninst) * 100, formatfloat
                ('0.00', (NsolInstPromethee[I] / Ninst) * 100) + '%');
            chrtresultadopromethee.BottomAxis.Items.Add(j * 2, inttostr(I + 1));
            // chrtResultado.Legend.Item[0].Text := inttostr(i + 1) + ' Solutions ';
            inc(j);
          end;
        end;
      end;

      chrtresultadopromethee.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.BottomAxis.visible := false;
      chrtresultadopromethee.Legend.Visible := false;
      chrtresultadopromethee.BottomAxis.LabelsAlternate;
      chrtresultadopromethee.BottomAxis.Items.Automatic := false;
      chrtresultadopromethee.LeftAxis.Maximum := 100;
      chrtresultadopromethee.LeftAxis.MaximumOffset := 16;
      chrtresultadopromethee.LeftAxis.Minimum := 0;
      // chrtresultado.LeftAxis.AutomaticMaximum := false;
      chrtresultadopromethee.LeftAxis.Increment := 20;
      chrtresultadopromethee.Title.Caption :=
        'Percentage of cases with X number of solutions';
      imgresultadopromethee.Picture.Bitmap.Assign
        (chrtresultadopromethee.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
            imgresultadopromethee.Width, imgresultadopromethee.Height)));
      { if numcrit <= 3 then
        grdresultadosPromethee.RowCount := 3;
        if numcrit > 3 then
        grdresultadosPromethee.RowCount := 4;
        if numcrit > 4 then
        grdresultadosPromethee.RowCount := 5;
        if numcrit > 6 then
        grdresultadosPromethee.RowCount := 6;




        // grdsegundo.RowCount := 6;

        grdresultados.ColumnCount := 2;
        // grdsegundo.ColumnCount := 1 + Numcrit;

        for j := 0 to grdresultados.ColumnCount - 1 do
        for I := 0 to grdresultados.RowCount - 1 do
        begin
        grdresultados.Cell[I, j].Alignment := taCenter;
        grdresultados.Cell[I, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
        for j := 0 to grdresultados.ColumnCount - 1 do
        grdresultados.Cell[0, j].Font.Style := [fsbold];

        for I := 0 to grdresultados.RowCount - 1 do
        grdresultados.Cell[I, 0].Font.Style := [fsbold];

        Fatorial := 1;
        for I := 1 to numcrit - 1 do
        Fatorial := Fatorial * I;
        // for i := 0 to Numcrit - 1 do
        // begin
        I := 0;
        grdresultados.Cell[0, I + 1].Text := 'Percentage';
        grdresultados.Cell[1, I + 1].Text := FormatFloat
        ('0.00', TotalPrimCrit * 100) + ' %';
        if numcrit > 3 then
        begin
        grdresultados.Cell[2, I + 1].Text := FormatFloat
        ('0.00', totalseccrit * 100) + ' %';
        grdresultados.Cell[3, I + 1].Text := FormatFloat('0.00', 1 * 100)
        + ' %';
        end
        else
        grdresultados.Cell[2, I + 1].Text := FormatFloat('0.00', 1 * 100)
        + ' %';
        if numcrit > 4 then
        begin
        grdresultados.Cell[3, I + 1].Text := FormatFloat
        ('0.00', totalterccrit * 100) + ' %';
        grdresultados.Cell[4, I + 1].Text := FormatFloat('0.00', 1 * 100)
        + ' %';
        end;
        if numcrit = 4 then
        grdresultados.Cell[3, I + 1].Text := FormatFloat('0.00', 1 * 100)
        + ' %';
        if numcrit > 6 then
        begin
        grdresultados.Cell[4, I + 1].Text := FormatFloat
        ('0.00', totalquartcrit * 100) + ' %';
        grdresultados.Cell[5, I + 1].Text := FormatFloat
        ('0.00', Fatorial / Fatorial * 100) + ' %';
        end;

        // end;

        grdresultados.Cell[0, 0].Text := 'Rank Position:';
        grdresultados.Cell[1, 0].Text := '1th:';
        grdresultados.Cell[2, 0].Text := 'Others:';
        if numcrit = 4 then
        begin
        grdresultados.Cell[2, 0].Text := '1th + 2th:';
        grdresultados.Cell[3, 0].Text := 'Others:';
        end;
        if (numcrit = 5) or (numcrit = 6) then
        begin
        grdresultados.Cell[2, 0].Text := '1th + 2th:';
        grdresultados.Cell[3, 0].Text := '1th + 2th + 3th:';
        grdresultados.Cell[4, 0].Text := 'Others:';
        end;
        if numcrit > 6 then
        begin
        grdresultados.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
        grdresultados.Cell[5, 0].Text := 'Others:';
        end; }

      grdnsolutionsPromethee.RowCount := 2;
      grdnsolutionsPromethee.ColumnCount := 2;

      for j := 0 to grdnsolutionsPromethee.ColumnCount - 1 do
        grdnsolutionsPromethee.Cell[0, j].Font.Style := [fsbold];

      // for i := 0 to grdresultados.RowCount - 1 do
      // grdnsolutions.Cell[i,0].Font.Style := [fsbold];

      grdnsolutionsPromethee.Cell[0, 0].Text := 'Number of Solutions';
      grdnsolutionsPromethee.Cell[0, 1].Text := 'Percentage';
      j := 0;
      for I := 0 to Length(NsolInstPromethee) - 1 do
      begin
        if round(NsolInstPromethee[I]) = 0 then
          continue
        else
        begin
          grdnsolutionsPromethee.RowCount := 2 + j;
          if I = 0 then
            grdnsolutionsPromethee.Cell[1 + j, 0].Text := '1 Solution'
          else
            grdnsolutionsPromethee.Cell[j + 1, 0].Text := inttostr(I + 1)
              + ' Solutions';
          grdnsolutionsPromethee.Cell[j + 1, 1].Text := formatfloat
            ('0.00', (NsolInstPromethee[I] / Ninst) * 100) + '%';
          inc(j);
        end;
      end;
      for j := 0 to grdnsolutionsPromethee.ColumnCount - 1 do
        for I := 0 to grdnsolutionsPromethee.RowCount - 1 do
        begin
          grdnsolutionsPromethee.Cell[I, j].Alignment := taCenter;
          grdnsolutionsPromethee.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      rgnmatcons.Visible := false;
      rgnresultadoPromethee.Visible := true;
      // btnshowresults.Enabled := true;

    end;
    //ExportarResultados;
  end;
end;

procedure TIWForm1.cmbbxauxiliarAsyncChange(Sender: TObject;
  EventParams: TStringList);
var
  I: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for I := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[I].ItemIndex = 2) or (cmbbxTipoCrit[I].ItemIndex = 3)
        then
        cmbbxNiveisCrit[I].Visible := true
      else
        cmbbxNiveisCrit[I].Visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxauxiliarChange(Sender: TObject);
var
  I: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for I := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[I].ItemIndex = 2) or (cmbbxTipoCrit[I].ItemIndex = 3)
        then
        cmbbxNiveisCrit[I].Visible := true
      else
        cmbbxNiveisCrit[I].Visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxmatrizconseqAltXSChange(Sender: TObject);
var
i, k, j : integer;
begin
with tiwusersession(webapplication.Data) do
begin
  if cmbbxmatrizconseqAltXS.ItemIndex = 0 then
  begin
  lnkmatrizconseqaltx.Visible := false;
    for k := 0 to 19 do
      begin
      chrtmatrizconseqAltXS.Series[k].Clear;
      chrtmatrizconseqAltXS.Series[k].Marks.FontSeriesColor := true;
      end;
      //chrtmatrizconseqAltXS.Series[1].Clear;


      //chrtmatrizconseqAltXS.Series[1].Marks.FontSeriesColor := true;

      chrtmatrizconseqAltXS.BottomAxis.Items.Clear;
      for i := 0 to numcrit - 1 do
      begin
      chrtmatrizconseqAltXS.BottomAxis.Items.Add(i,'Criterion ' + inttostr(I+1));
      end;
      for k := 0 to 9 do
      begin
      for i := 0 to numcrit - 1 do
      begin
      chrtmatrizconseqAltXS.Series[k].AddXY(i, Tiwusersession(webapplication.data).matrizconseqaltxOrdenado[ninst-1 -10*k,i]);

      chrtmatrizconseqAltXS.Series[k+10].AddXY(i, Tiwusersession(webapplication.data).matrizconseqaltSOrdenado[ninst-1 -10*k,i]);
      end;
      end;

      imgmatrizconseqAltXS.Picture.Bitmap.Assign
        (chrtmatrizconseqAltXS.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgmatrizconseqAltXS.Width,
            imgmatrizconseqAltXS.Height)));
  end;
  if cmbbxmatrizconseqAltXS.ItemIndex > 0 then
  begin
  lnkmatrizconseqaltx.Visible := true;
    chrtpercentilMatrizConseqXS.Series[0].Clear;
    chrtpercentilMatrizConseqXS.Series[1].Clear;
    chrtpercentilMatrizConseqXS.Series[0].Marks.FontSeriesColor := true;
    chrtpercentilMatrizConseqXS.Series[1].Marks.FontSeriesColor := true;
    chrtpercentilMatrizConseqXS.BottomAxis.Items.Clear;
      for i := 0 to numcrit - 1 do
      for j := 0 to numcrit - 1 do
      begin
      if CasesOrdemcrit[OrdemPerdaMaximaX[10*cmbbxmatrizconseqAltXS.ItemIndex-1],j] = i+1 then
      begin
      chrtpercentilMatrizConseqXS.BottomAxis.Items.Add(i,'Criterion ' + inttostr(j+1));
      chrtpercentilMatrizConseqXS.Series[0].AddXY(i, Tiwusersession(webapplication.data).matrizconseqaltx[10*cmbbxmatrizconseqAltXS.ItemIndex-1,j]);
      chrtpercentilMatrizConseqXS.Series[1].AddXY(i, Tiwusersession(webapplication.data).matrizconseqaltS[10*cmbbxmatrizconseqAltXS.ItemIndex-1,j]);
      end;
      end;

     imgmatrizconseqAltXS.Picture.Bitmap.Assign
        (chrtpercentilMatrizConseqXS.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgmatrizconseqAltXS.Width,
            imgmatrizconseqAltXS.Height)));
  end;

end;
end;

procedure TIWForm1.cmbbxprobaltChange(Sender: TObject);
begin
  if cmbbxprobalt.ItemIndex = 0 then
    imgresultadosprobabilidade.Picture.Bitmap.Assign
      (chrtprobalt1.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgresultadosprobabilidade.Width, imgresultadosprobabilidade.Height))
      );
  if cmbbxprobalt.ItemIndex = 1 then
    imgresultadosprobabilidade.Picture.Bitmap.Assign
      (chrtprobalt2.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgresultadosprobabilidade.Width, imgresultadosprobabilidade.Height))
      );
  if cmbbxprobalt.ItemIndex = 2 then
    imgresultadosprobabilidade.Picture.Bitmap.Assign
      (chrtprobalt3.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgresultadosprobabilidade.Width, imgresultadosprobabilidade.Height))
      );

end;

procedure TIWForm1.rgnalternativaSCreate(Sender: TObject);
begin
rgnalternativaS.Align := AlClient;
end;

procedure TIWForm1.rgnalternativawCreate(Sender: TObject);
begin
rgnalternativaw.Align := AlClient;
end;

procedure TIWForm1.rgnalternativaYCreate(Sender: TObject);
begin
rgnalternativaY.Align := alclient;
end;

procedure TIWForm1.rgnmatconsCreate(Sender: TObject);
begin
  rgnmatcons.Align := AlClient;
end;

procedure TIWForm1.rgnmatrizconseqAltXSCreate(Sender: TObject);
begin
rgnmatrizconseqAltXS.Align := alclient;
end;

procedure TIWForm1.rgnperdamaximaCreate(Sender: TObject);
begin
rgnperdamaxima.Align := alclient;
end;

procedure TIWForm1.rgnalternativaZCreate(Sender: TObject);
begin
rgnalternativaZ.Align := AlClient;
end;

procedure TIWForm1.rgnConsequenceRangeCreate(Sender: TObject);
begin
rgnconsequencerange.Align := AlClient;
end;

procedure TIWForm1.rgndifvgCreate(Sender: TObject);
begin
rgndifvg.Align := alclient;
end;

procedure TIWForm1.rgnregraCreate(Sender: TObject);
begin
  rgnregra.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoCreate(Sender: TObject);
begin
  rgnresultado.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoPrometheeCreate(Sender: TObject);
begin
  rgnresultadoPromethee.Align := AlClient;
end;

procedure TIWForm1.rgnresultadosprobabilidadeCreate(Sender: TObject);
begin
  rgnresultadosprobabilidade.Align := AlClient;
end;

initialization

TIWForm1.SetAsMainForm;

end.
