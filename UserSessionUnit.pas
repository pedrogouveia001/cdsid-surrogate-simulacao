unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompLabel,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, TeeGDIPlus, IWCompButton, TeEngine, ExtCtrls,
  TeeProcs, Chart, IWExtCtrls, IWWebGrid, IWAdvWebGrid, BubbleCh, TeePolar,
  IWCompListbox, IWCompRadioButton, IWTMSCheckList, series, IWGrids, TeeLisB,
  TeeSeriesTextEd, Buttons, Math, pngimage, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns, ActnList, ImgList,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Windows, Messages, Variants, Graphics, Dialogs, ComCtrls, Menus, StdCtrls,
  jpeg, ComObj, Grids, IWCompText, IWHTMLControls, ErrorBar,
  IWAdvWebGridExcel, IWProgressIndicator, IWDBGrids,
  IWCompProgressBar, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWCompTabControl, IWHelpTip, IWTMSPopup, IWTMSMenus, GanttCh, IWCompMemo,
  IWCompCheckbox;
type
  TIWUserSession = class(TIWUserSessionBase)
  private
    { Private declarations }
  public
    { Public declarations }
    edtNomeCrit : array of TiwEdit;
    edtNomeAlt : array of TiwEdit;
    cmbbxTipoCrit : array of TiwComboBox;
    cmbbxNiveisCrit : array of TiwComboBox;
    edtMatrizConseq : array of array of TiwEdit;

    NomeCrit, NomeAlt : array of string;
    NumCrit, Numalt : integer;
    Tipocrit : array of integer;
    Niveis : array of integer;
    MatrizConseq, MatrizConseqNorm : array of array of double;
    PesoCrit : array of array of double;
    CasesOrdemCrit : array of array of integer;
    //OrdemCrit : array of integer;
    Max, Min : array of double;
    MatrizPOA, MatrizPOAPromethee : array of array of integer;
    ResultadoROC, ResultadoPromethee : array of array of double;
    MatrizSol, MatrizSolPromethee : array of array of double;
    ResultAlt, ResultaltPromethee : array of double;
    ResultSol, ResultSolPromethee : array of double;
    SobClassMatrix: array of array of real;
    PositiveFlow, NegativeFlow, NetFlow : array of real;
    matrizParaPar: array of array of array of real;
    comp1, comp2, ind, pref, gauss: real;
    coresalt, coresalt1: array of Tcolor;
    CaseSol, CasesolPromethee : array of integer;
    excelverif : boolean;
    Total: double;
    NsolInst, NsolInstPromethee : array of integer;
    Ninst : integer;
    PrimCrit, SecCrit, TercCrit, QuartCrit : array of real;
    TotalPrimCrit, totalseccrit, totalterccrit, totalquartcrit : real;
    QtdMedia, DesvioPadrao, mediana, moda : double;
    QtdMediaPromethee, DesvioPadraoPromethee, medianaPromethee, modaPromethee : double;
    auxpromethee : integer;
    Aditivo, Sobreclassificacao : boolean;
    MatrizDifVG, MatrizDifVGPromethee : array of array of double;
    PerdaMedia, PerdaMaxima, probalt : array of double;
    PerdaMediaPromethee, PerdaMaximaPromethee, probaltPromethee : array of double;
    MediaPerdaMedia, MediaPerdaMediaPromethee : array [0..3] of array [0..4] of double;
    MaximaPerdaMedia, MaximaPerdaMediaPromethee : array [0..3] of array [0..4] of double;
    MediaPerdaMaxima, MediaPerdaMaximaPromethee : array [0..3] of array [0..4] of double;
    MaximaPerdaMaxima, MaximaPerdaMaximaPromethee : array [0..3] of array [0..4] of double;
    Regra1, Regra2, Regra3, regra4 : integer;
    Regra1Promethee, Regra2Promethee, Regra3Promethee, regra4Promethee : integer;
    probregra, probregrapromethee : array [0..3] of array [0..4] of integer;
    percentualr1, percentualr2, percentualr3, percentualr4 : double;
    percentualr1Promethee, percentualr2Promethee, percentualr3Promethee, percentualr4Promethee : double;
    perdamediaindividuos, perdamaximaindividuos : array of double;
    perdamediapopulacao, perdamaximapopulacao : array of double;
    problimiarperdamedia, problimiarperdamaxima : array of double;
    desempenhoindividuos : array of double;
    mediaprobalt, maximoprobalt, desviopadraoprobalt : double;
    probaltestratificado, probaltestratificado2, probaltestratificado3  : array [0..100] of double;
    matrizprobalt : array of double;
    instancia : integer;
    medianaprobalt, modaprobalt : double;
    faixa40probalt, faixa50probalt, faixa60probalt, faixa70probalt, faixa80probalt, faixa90probalt : double;
    faixa40probaltacumulado, faixa50probaltacumulado, faixa60probaltacumulado, faixa70probaltacumulado, faixa80probaltacumulado, faixa90probaltacumulado : double;
    faixa40probalt2, faixa50probalt2, faixa60probalt2, faixa70probalt2, faixa80probalt2, faixa90probalt2 : double;
    faixa40probaltacumulado2, faixa50probaltacumulado2, faixa60probaltacumulado2, faixa70probaltacumulado2, faixa80probaltacumulado2, faixa90probaltacumulado2 : double;
    faixa40probalt3, faixa50probalt3, faixa60probalt3, faixa70probalt3, faixa80probalt3, faixa90probalt3 : double;
    faixa40probaltacumulado3, faixa50probaltacumulado3, faixa60probaltacumulado3, faixa70probaltacumulado3, faixa80probaltacumulado3, faixa90probaltacumulado3 : double;

    ganhomedioindividuos : array of double;

    desempenhopopulacao : array of double;
    mediadesempenho, maximodesempenho : double;
    percentualdesempenho : array[0..200] of double;

    mediaprobalt2, mediaprobalt3, desviopadraoprobalt2, desviopadraoprobalt3  : double;

    MatrizRankingAlt : array of array of integer;
    RankingPrimeira, RankingSegunda, RankingTerceira : array of double;
    MatrizFaixaPrimeira, MatrizFaixaSegunda, MatrizFaixaTerceira : array of array of double;
    VetorAlternativaX, VetorAlternativaZ, VetorAlternativaS : array of array of double;
    QuartisAlternativaX, QuartisAlternativaZ, QuartisAlternativaS : array of array of double;
    PercentilAlternativaX : array of array of double; // Pega a probabilidade associada ao percentil e a posição do ranking da melhor alternativa.
    PercentilAlternativaZ : array of array of double; // Pega a probabilidade associada ao percentil e a posição do ranking da segunda melhor alternativa.
    PercentilAlternativaS : array of array of double; // Pega a probabilidade associada ao percentil e a posição do ranking da alt com maior prob de ser segunda
    PercentualSigualaX : double;
    VetorAlternativaY : array of array of double;
    QuartisAlternativaY : array of array of double;
    PercentilAlternativaY : array of array of double;
    VetorAlternativaW : array of array of double;
    QuartisAlternativaW : array of array of double;
    PercentilAlternativaW : array of array of double;
    NomeProblema : string;

    PerdaMaximaX, PerdaMaximaZ, PerdaMaximaS, PerdaMaximaW : array of double;
    QuartisPerdaMaximaX, QuartisPerdaMaximaZ, QuartisPerdaMaximaS, QuartisPerdaMaximaW : array of double;
    PercentilPerdaMaximaX, PercentilPerdaMaximaZ, PercentilPerdaMaximaS, PercentilPerdaMaximaW : array of double;


    ConsequenceRangeX, ConsequenceRangeS, ConsequenceRangeZ : array of double;
    quartisConsequenceRangeX, quartisConsequenceRangeS, quartisConsequenceRangeZ : array of double;
    percentilConsequenceRangeX, percentilConsequenceRangeS, percentilConsequenceRangeZ : array of double;
    difVGX, difVGZ, difVGS : array of double;
    quartisdifVGX, quartisDifVGZ, quartisdifVGS : array of double;
    percentildifVGX, percentildifVGZ, percentildifVGS : array of double;
    MatrizConseqAltX, MatrizConseqAltS, MatrizConseqAltW : array of array of double;
    MatrizConseqAltXOrdenado, MatrizConseqAltSOrdenado, MatrizConseqAltWOrdenado : array of array of double;
    OrdemPerdaMaximaX, ordemperdamaximas, ordemperdamaximaz, OrdemperdamaximaW : array of integer;

    PercentualPerdaMaximaXPesosIguais : double;
    PercentualPerdaMaximaSPesosIguais : double;
    PercentualPerdaMaximaZPesosIguais : double;

    MatrizConseqPercentilPerdaX : array of array of array of double;
    //MatrizConseqPercentilPerdaS : array of array of array of double;
  end;

implementation

uses UntEntrada;

{$R *.dfm}

end.