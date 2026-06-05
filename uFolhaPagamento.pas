unit uFolhaPagamento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Grids, DBGrids, DB, DBClient;

type
  TfrFolhaPagamento = class(TForm)
    gbFuncionario: TGroupBox;
    lbNome: TLabel;
    cbNome: TComboBox;
    lbCargo: TLabel;
    edCargo: TEdit;
    btCadastrar: TButton;
    pnCadastroFuncionario: TPanel;
    lbCadastroFuncionario: TLabel;
    lbCodigoFuncionario: TLabel;
    edCodigoFuncionario: TEdit;
    lbNomeFuncionario: TLabel;
    edNomeFuncionario: TEdit;
    Cargo: TLabel;
    cbCargo: TComboBox;
    lbEndereco: TLabel;
    edEndereco: TEdit;
    lbTelefone: TLabel;
    edTelefone: TEdit;
    btFechar: TButton;
    btSalvarFuncionario: TButton;
    gbProventos: TGroupBox;
    lbSalarioBase: TLabel;
    edSalarioBase: TEdit;
    lbHorasExtras: TLabel;
    edHorasExtras: TEdit;
    lbOutros: TLabel;
    edOutros: TEdit;
    edTotalProventos: TEdit;
    lbTotal: TLabel;
    gbDescontos: TGroupBox;
    lbINSS: TLabel;
    lbIrrf: TLabel;
    lbValeTransporte: TLabel;
    lbTotalDescontos: TLabel;
    edInss: TEdit;
    edIrrf: TEdit;
    edValeTransporte: TEdit;
    edTotalDescontos: TEdit;
    gbResultado: TGroupBox;
    edResultadoTotalProventos: TEdit;
    lbResultadoTotalProventos: TLabel;
    edResultadoTotalDescontos: TEdit;
    lbResultadoTotalDescontos: TLabel;
    btLimpar: TButton;
    btSalvarFolha: TButton;
    btCalcular: TButton;
    grFolhaPagamento: TDBGrid;
    lbSalarioLiquido: TLabel;
    edSalarioLiquido: TEdit;
    gbIdentificacao: TGroupBox;
    dsFolhaPagamento: TDataSource;
    btConsultar: TButton;
    pnConsultaFuncionarios: TPanel;
    lbTituloFuncionario: TLabel;
    btFecharConsultaFuncionario: TButton;
    cdsFuncionario: TClientDataSet;
    cdsFuncionariobdCODIGOFUNCIONARIO: TIntegerField;
    cdsFuncionariobdNOMEFUNCIONARIO: TStringField;
    cdsFuncionariobdCARGO: TStringField;
    cdsFuncionariobdENDERECO: TStringField;
    cdsFuncionariobdTELEFONE: TStringField;
    dsFuncionarios: TDataSource;
    grFuncionarios: TDBGrid;
    btBuscar: TButton;
    lbBuscaNome: TLabel;
    edBuscaNome: TEdit;
    cdsFolhaPagamento: TClientDataSet;
    cdsFolhaPagamentobdCODIGO: TIntegerField;
    cdsFolhaPagamentobdCODIGOFUNCIONARIO: TIntegerField;
    cdsFolhaPagamentobdNOME: TStringField;
    cdsFolhaPagamentobdCARGO: TStringField;
    cdsFolhaPagamentobdSALARIOBASE: TCurrencyField;
    cdsFolhaPagamentobdHORASEXTRAS: TCurrencyField;
    cdsFolhaPagamentobdOUTROS: TCurrencyField;
    cdsFolhaPagamentobdTOTALPROVENTOS: TCurrencyField;
    cdsFolhaPagamentobdINSS: TCurrencyField;
    cdsFolhaPagamentobdIRRF: TCurrencyField;
    cdsFolhaPagamentobdVALETRANSPORTE: TCurrencyField;
    cdsFolhaPagamentobdTOTALDESCONTOS: TCurrencyField;
    cdsFolhaPagamentobdSALARIOLIQUIDO: TCurrencyField;
    cdsFolhaPagamentobdANO: TIntegerField;
    cdsFolhaPagamentobdMES: TStringField;
    lbMes: TLabel;
    cbMes: TComboBox;
    edAno: TEdit;
    lbAno: TLabel;
    GroupBox1: TGroupBox;
    edCodigo: TEdit;
    lbCodigo: TLabel;
    btDeletarFuncionario: TButton;
    btDeletarFolha: TButton;
    procedure btCadastrarClick(Sender: TObject);
    procedure btFecharClick(Sender: TObject);
    procedure btSalvarFuncionarioClick(Sender: TObject);
    procedure cbNomeSelect(Sender: TObject);
    procedure edSalarioBaseExit(Sender: TObject);
    procedure edHorasExtrasExit(Sender: TObject);
    procedure edOutrosExit(Sender: TObject);
    procedure btCalcularClick(Sender: TObject);
    procedure btSalvarFolhaClick(Sender: TObject);
    procedure btLimparClick(Sender: TObject);
    procedure btConsultarClick(Sender: TObject);
    procedure btFecharConsultaFuncionarioClick(Sender: TObject);
    procedure edSalarioBaseEnter(Sender: TObject);
    procedure edHorasExtrasEnter(Sender: TObject);
    procedure edOutrosEnter(Sender: TObject);
    procedure edCodigoExit(Sender: TObject);
    procedure btBuscarClick(Sender: TObject);
    procedure grFuncionariosCellClick(Column: TColumn);
    procedure grFolhaPagamentoCellClick(Column: TColumn);
    procedure edCodigoFuncionarioExit(Sender: TObject);
    procedure edCodigoEnter(Sender: TObject);
    procedure btDeletarFuncionarioClick(Sender: TObject);
    procedure btDeletarFolhaClick(Sender: TObject);
  private
    { Private declarations }

    // Variáveis Funcionário
    wCodigoFuncionario: Integer;
    wNomeFuncionario: String;
    wCargoFuncionario: String;
    wEnderecoFuncionario: String;
    wTelefoneFuncionario: String;

    // Variáveis Folha de Pagamento
    wCodigoFolha: Integer;
    wMesFolha: String;
    wAnoFolha: Integer;
    wNomeFolha: String;
    wCargoFolha: String;

    // Proventos
    wSalarioBase: Currency;
    wHorasExtras: Currency;
    wOutros: Currency;
    wTotalProventos: Currency;

    // Descontos
    wInss: Currency;
    wIrrf: Currency;
    wValeTransporte: Currency;
    wTotalDescontos: Currency;

    // Resultado
    wSalarioLiquido: Currency;

    // Tela
    wConsultaAtiva: Boolean;

    procedure pCadastrarFuncionario;
    procedure pEditarFuncionario;
    procedure pBuscarFuncionario;
    procedure pDeletarFuncionario;
    procedure pSalvarFolha;
    procedure pBuscarFolha;
    procedure pCarregarFolha;
    procedure pDeletarFolha;
    procedure pCalcularTotalProventos;
    procedure pCalcularTotalDescontos;
    procedure pExibirCadastro;
    procedure pModoCadastro;
    procedure pModoConsulta;
    procedure pAbrirConsulta;
    procedure pLimparCampos;
    procedure pLimparCamposCadastroFuncionario;

    // Validacoes
    function fValidarCamposFuncionario: Boolean;
    function fValidarCamposFolha: Boolean;
    function fValidarCampoSalarioBase: Boolean;

    // INSS

    function fCalcularInss: Currency;
    function fCalcularPrimeiraFaixa: Currency;
    function fCalcularSegundaFaixa: Currency;
    function fCalcularTerceiraFaixa: Currency;

  public
    { Public declarations }
  end;

var
  frFolhaPagamento: TfrFolhaPagamento;

implementation

uses Math;

{$R *.dfm}

// Botoes

procedure TfrFolhaPagamento.btCadastrarClick(Sender: TObject);
begin
  pExibirCadastro;
  pModoCadastro;
end;

procedure TfrFolhaPagamento.btSalvarFuncionarioClick(Sender: TObject);

begin
  wCodigoFuncionario     := StrToIntDef(edCodigoFuncionario.Text,0);
  wNomeFuncionario       := edNomeFuncionario.Text;
  wCargoFuncionario      := cbCargo.Text;
  wEnderecoFuncionario   := edEndereco.Text;
  wTelefoneFuncionario   := edTelefone.Text;

  if not fValidarCamposFuncionario then
     Exit;

  cdsFuncionario.IndexFieldNames := 'bdCODIGOFUNCIONARIO';

  if not cdsFuncionario.FindKey([edCodigoFuncionario.Text]) then
     begin
       pCadastrarFuncionario;
     end
  else
     begin
       pEditarFuncionario;
     end;
end;

procedure TfrFolhaPagamento.btConsultarClick(Sender: TObject);
begin
  pnConsultaFuncionarios.Visible := True;
  pnConsultaFuncionarios.Top  := 100;
  pnConsultaFuncionarios.Left := 100;
end;

procedure TfrFolhaPagamento.btBuscarClick(Sender: TObject);
begin
  cdsFuncionario.IndexFieldNames := 'bdNOMEFUNCIONARIO';

  if not cdsFuncionario.FindKey([edBuscaNome.Text]) then
    begin
      ShowMessage('Funcionário não encontrado');
      Exit;
    end;
end;

procedure TfrFolhaPagamento.btCalcularClick(Sender: TObject);
begin
  if not fValidarCampoSalarioBase then
     Exit;

  pCalcularTotalDescontos;

  edInss.Text                    := CurrToStr(wInss);
  edIrrf.Text                    := CurrToStr(wIrrf);
  edValeTransporte.Text          := CurrToStr(wValeTransporte);
  edTotalDescontos.Text          := CurrToStr(wTotalDescontos);
  edResultadoTotalProventos.Text := CurrToStr(wTotalProventos);
  edResultadoTotalDescontos.Text := CurrToStr(wTotalDescontos);
  edSalarioLiquido.Text          := CurrToStr(wSalarioLiquido);
  btSalvarFolha.Enabled := True;
end;

procedure TfrFolhaPagamento.btSalvarFolhaClick(Sender: TObject);
begin
  if (not fValidarCamposFolha) or (not fValidarCampoSalarioBase) then
     Exit;

  pSalvarFolha;

  pLimparCampos;
  cbNome.SetFocus;

  btSalvarFolha.Enabled := False;
end;

procedure TfrFolhaPagamento.btLimparClick(Sender: TObject);
begin
  pLimparCampos;
  cbNome.SetFocus;
end;

procedure TfrFolhaPagamento.btFecharClick(Sender: TObject);
begin
  pnCadastroFuncionario.Visible := False;
  pnCadastroFuncionario.Left := 648;

  pLimparCamposCadastroFuncionario;

  if wConsultaAtiva then
     pAbrirConsulta;

  wConsultaAtiva := False;
end;

procedure TfrFolhaPagamento.btFecharConsultaFuncionarioClick(
  Sender: TObject);
begin
  pnConsultaFuncionarios.Visible := False;
  pnConsultaFuncionarios.Top := 280;
  pnConsultaFuncionarios.Left := 680;
end;

// Eventos dos componentes

procedure TfrFolhaPagamento.cbNomeSelect(Sender: TObject);
begin
  cdsFuncionario.IndexFieldNames := 'bdNOMEFUNCIONARIO';

  if cdsFuncionario.FindKey([cbNome.Text]) then
     begin
       edCargo.Text := cdsFuncionariobdCARGO.AsString;
     end
end;

procedure TfrFolhaPagamento.edCodigoFuncionarioExit(Sender: TObject);
begin
  if edCodigoFuncionario.Text <> '' then
     pBuscarFuncionario;
end;

procedure TfrFolhaPagamento.edSalarioBaseEnter(Sender: TObject);
begin
  if edSalarioBase.Text = '0,00' then
     begin
       edSalarioBase.Text := '';
     end;
end;

procedure TfrFolhaPagamento.edHorasExtrasEnter(Sender: TObject);
begin
   if edHorasExtras.Text = '0,00' then
      begin
        edHorasExtras.Text := '';
      end;
end;

procedure TfrFolhaPagamento.edOutrosEnter(Sender: TObject);
begin
    if edOutros.Text = '0,00' then
       begin
         edOutros.Text := '';
       end;
end;

procedure TfrFolhaPagamento.edSalarioBaseExit(Sender: TObject);
begin
  pCalcularTotalProventos;
end;

procedure TfrFolhaPagamento.edHorasExtrasExit(Sender: TObject);
begin
  pCalcularTotalProventos;
end;

procedure TfrFolhaPagamento.edOutrosExit(Sender: TObject);
begin
  pCalcularTotalProventos;
end;

procedure TfrFolhaPagamento.edCodigoExit(Sender: TObject);
begin
  if edCodigo.Text <> '' then
     pBuscarFolha;
end;

// Grids

procedure TfrFolhaPagamento.grFuncionariosCellClick(Column: TColumn);
begin
  edCodigoFuncionario.Text := IntToStr(cdsFuncionariobdCODIGOFUNCIONARIO.AsInteger);
  edNomeFuncionario.Text   := cdsFuncionariobdNOMEFUNCIONARIO.AsString;
  edEndereco.Text          := cdsFuncionariobdENDERECO.AsString;
  edTelefone.Text          := cdsFuncionariobdTELEFONE.AsString;
  cbCargo.ItemIndex        := cbCargo.Items.IndexOf(cdsFuncionariobdCARGO.AsString);

  pnConsultaFuncionarios.Visible := False;

  pExibirCadastro;

  pModoConsulta;
end;

procedure TfrFolhaPagamento.grFolhaPagamentoCellClick(Column: TColumn);
begin
  edCodigo.Text := IntToStr(cdsFolhaPagamentobdCODIGO.AsInteger);
  pCarregarFolha;
end;

// Procedures principais

procedure TfrFolhaPagamento.pCadastrarFuncionario;
begin
   cdsFuncionario.Insert;

   cdsFuncionariobdCODIGOFUNCIONARIO.AsInteger := wCodigoFuncionario;
   cdsFuncionariobdNOMEFUNCIONARIO.AsString    := wNomeFuncionario;
   cdsFuncionariobdCARGO.AsString              := wCargoFuncionario;
   cdsFuncionariobdENDERECO.AsString           := wEnderecoFuncionario;
   cdsFuncionariobdTELEFONE.AsString           := wTelefoneFuncionario;

   cdsFuncionario.Post;

   cbNome.Items.Add(wNomeFuncionario);

   ShowMessage('Cadastro realizado com sucesso');

   pLimparCamposCadastroFuncionario;
end;

procedure TfrFolhaPagamento.pEditarFuncionario;
var
  wIndexFuncionario : Integer;
begin
  wIndexFuncionario := cbNome.Items.IndexOf(cdsFuncionariobdNOMEFUNCIONARIO.AsString);

  cbNome.Items.Delete(wIndexFuncionario);

  cdsFuncionario.Edit;

  cdsFuncionariobdCODIGOFUNCIONARIO.AsInteger := wCodigoFuncionario;
  cdsFuncionariobdNOMEFUNCIONARIO.AsString    := wNomeFuncionario;
  cdsFuncionariobdCARGO.AsString              := wCargoFuncionario;
  cdsFuncionariobdENDERECO.AsString           := wEnderecoFuncionario;
  cdsFuncionariobdTELEFONE.AsString           := wTelefoneFuncionario;

  cdsFuncionario.Post;

  cbNome.Items.Add(wNomeFuncionario);

  cdsFolhaPagamento.IndexFieldNames := 'bdCODIGOFUNCIONARIO';

  cdsFolhaPagamento.First;

  while not cdsFolhaPagamento.Eof do
    begin
       if cdsFolhaPagamentobdCODIGOFUNCIONARIO.AsInteger = wCodigoFuncionario then
          begin
            cdsFolhaPagamento.Edit;

            cdsFolhaPagamentobdNOME.AsString := wNomeFuncionario;

            cdsFolhaPagamento.Post;
          end;

       cdsFolhaPagamento.Next;
    end;
    
   ShowMessage('Cadastro Editado com Sucesso');

   pLimparCamposCadastroFuncionario;
end;

procedure TfrFolhaPagamento.pBuscarFuncionario;
var
  wCodigoBusca: Integer;
begin
  cdsFuncionario.IndexFieldNames := 'bdCODIGOFUNCIONARIO';

  wCodigoBusca := StrToIntDef(edCodigoFuncionario.Text, 0);

  if wCodigoBusca = 0 then
     begin
       ShowMessage('Informe um código válido');
       edCodigoFuncionario.Clear;
       edCodigoFuncionario.SetFocus;
       Exit;
     end;

  if cdsFuncionario.FindKey([wCodigoBusca]) then
     begin
       edNomeFuncionario.Text := cdsFuncionariobdNOMEFUNCIONARIO.AsString;
       cbCargo.ItemIndex      := cbCargo.Items.IndexOf(cdsFuncionariobdCARGO.AsString);
       edEndereco.Text        := cdsFuncionariobdENDERECO.AsString;
       edTelefone.Text        := cdsFuncionariobdTELEFONE.AsString;

       btDeletarFuncionario.Enabled := True;
     end;
end;

procedure TfrFolhaPagamento.pSalvarFolha;
begin
  cdsFolhaPagamento.IndexFieldNames := 'bdCODIGO';

  if not cdsFolhaPagamento.FindKey([edCodigo.Text]) then
     begin
       cdsFolhaPagamento.Last;

       if cdsFolhaPagamento.RecordCount = 0 then
          wCodigoFolha := 1
       else
          wCodigoFolha := cdsFolhaPagamentobdCODIGO.AsInteger + 1;
          cdsFolhaPagamento.Insert;
     end
  else
     begin
       wCodigoFolha := cdsFolhaPagamentobdCODIGO.AsInteger;
       cdsFolhaPagamento.Edit;
     end;

  cdsFolhaPagamentobdCODIGO.AsInteger            := wCodigoFolha;
  cdsFolhaPagamentobdCODIGOFUNCIONARIO.AsInteger := cdsFuncionariobdCODIGOFUNCIONARIO.AsInteger;
  cdsFolhaPagamentobdMES.AsString                := wMesFolha;
  cdsFolhaPagamentobdANO.AsInteger               := wAnoFolha;
  cdsFolhaPagamentobdNOME.AsString               := wNomeFolha;
  cdsFolhaPagamentobdCARGO.AsString              := wCargoFolha;
  cdsFolhaPagamentobdSALARIOBASE.AsCurrency      := wSalarioBase;
  cdsFolhaPagamentobdHORASEXTRAS.AsCurrency      := wHorasExtras;
  cdsFolhaPagamentobdOUTROS.AsCurrency           := wOutros;
  cdsFolhaPagamentobdTOTALPROVENTOS.AsCurrency   := wTotalProventos;
  cdsFolhaPagamentobdINSS.AsCurrency             := wInss;
  cdsFolhaPagamentobdIRRF.AsCurrency             := wIrrf;
  cdsFolhaPagamentobdVALETRANSPORTE.AsCurrency   := wValeTransporte;
  cdsFolhaPagamentobdTOTALDESCONTOS.AsCurrency   := wTotalDescontos;
  cdsFolhaPagamentobdSALARIOLIQUIDO.AsCurrency   := wSalarioLiquido;

  cdsFolhaPagamento.Post;
end;

procedure TfrFolhaPagamento.pBuscarFolha;
var
  wCodigoBusca: Integer;
begin
  wCodigoBusca := StrToIntDef(edCodigo.Text,0);

  if wCodigoBusca = 0 then
     begin
       ShowMessage('Informe um código válido');
       edCodigo.Clear;
       edCodigo.SetFocus;
       Exit;
     end;
     
  pCarregarFolha;
end;

procedure TfrFolhaPagamento.pCarregarFolha;
begin
  cdsFolhaPagamento.IndexFieldNames := 'bdCODIGO';

  if cdsFolhaPagamento.FindKey([edCodigo.Text]) then
     begin
       btSalvarFolha.Enabled := False;
       cbNome.Enabled  := False;

       cbMes.ItemIndex                := cbMes.Items.IndexOf(cdsFolhaPagamentobdMES.AsString);
       edAno.Text                     := cdsFolhaPagamentobdANO.AsString;
       cbNome.ItemIndex               := cbNome.Items.IndexOf(cdsFolhaPagamentobdNOME.AsString);
       edCargo.Text 	    	          := cdsFolhaPagamentobdCARGO.AsString;
       edSalarioBase.Text 	          := CurrToStr(cdsFolhaPagamentobdSALARIOBASE.AsCurrency);
       edHorasExtras.Text 	          := CurrToStr(cdsFolhaPagamentobdHORASEXTRAS.AsCurrency);
       edOutros.Text	    	 	        := CurrToStr(cdsFolhaPagamentobdOUTROS.AsCurrency);
       edINSS.Text 		 		            := CurrToStr(cdsFolhaPagamentobdINSS.AsCurrency);
       edIRRF.Text 		 		            := CurrToStr(cdsFolhaPagamentobdIRRF.AsCurrency);
       edValeTransporte.Text 		      := CurrToStr(cdsFolhaPagamentobdVALETRANSPORTE.AsCurrency);
       edTotalProventos.Text  		    := CurrToStr(cdsFolhaPagamentobdTOTALPROVENTOS.AsCurrency);
       edTotalDescontos.Text 		      := CurrToStr(cdsFolhaPagamentobdTOTALDESCONTOS.AsCurrency);
       edResultadoTotalProventos.Text := CurrToStr(cdsFolhaPagamentobdTOTALPROVENTOS.AsCurrency);
       edResultadoTotalDescontos.Text := CurrToStr(cdsFolhaPagamentobdTOTALDESCONTOS.AsCurrency);
       edSalarioLiquido.Text          := CurrToStr(cdsFolhaPagamentobdSALARIOLIQUIDO.AsCurrency);
       wSalarioBase                   := cdsFolhaPagamentobdSALARIOBASE.AsCurrency;

       btDeletarFolha.Enabled := True;
     end
  else
    begin
      ShowMessage('Folha não encontrada');
    end;
end;

procedure TfrFolhaPagamento.pCalcularTotalProventos;
begin
  wSalarioBase := StrToCurrDef(edSalarioBase.Text,0);
  wHorasExtras := StrToCurrDef(edHorasExtras.Text,0);
  wOutros      := StrToCurrDef(edOutros.Text,0);

  wTotalProventos := wSalarioBase + wHorasExtras + wOutros;

  if wSalarioBase = 0 then
     begin
       edSalarioBase.Text := '0,00';
     end;

  if wHorasExtras = 0 then
     begin
       edHorasExtras.Text := '0,00';
     end;

  if wOutros = 0 then
     begin
       edOutros.Text := '0,00';
     end;

  if wTotalProventos = 0 then
     edTotalProventos.Text := '0,00'
  else
     edTotalProventos.Text := CurrToStr(wTotalProventos);
end;

procedure TfrFolhaPagamento.pCalcularTotalDescontos;
begin
  wInss := fCalcularInss;
  wIrrf := wSalarioBase * 0.15;
  wValeTransporte := wSalarioBase * 0.06;

  wTotalDescontos := wInss + wIrrf + wValeTransporte;
  wSalarioLiquido := wTotalProventos - wTotalDescontos;
end;

// Inteface

procedure TfrFolhaPagamento.pExibirCadastro;
begin
  pnCadastroFuncionario.Visible  := True;
  pnCadastroFuncionario.Left := 130;
  pnCadastroFuncionario.Top := 100;
end;

procedure TfrFolhaPagamento.pModoCadastro;
begin
  lbCadastroFuncionario.Caption := 'Cadastro de funcionário';
  btDeletarFuncionario.Enabled := False;
end;

procedure TfrFolhaPagamento.pModoConsulta;
begin
  lbCadastroFuncionario.Caption := 'Consulta de funcionário';
  wConsultaAtiva := True;
  btDeletarFuncionario.Enabled := True;
end;

procedure TfrFolhaPagamento.pLimparCampos;
begin
  cbMes.ItemIndex := -1;
  edAno.Clear;
  cbNome.ItemIndex := -1;
  edCargo.Clear;
  edSalarioBase.Text := '0,00';
  edHorasExtras.Text := '0,00';
  edOutros.Text := '0,00';
  edTotalProventos.Text := '0,00';
  edInss.Text := '0,00';
  edIrrf.Text := '0,00';
  edValeTransporte.Text := '0,00';
  edTotalDescontos.Text := '0,00';
  edResultadoTotalProventos.Text := '0,00';
  edResultadoTotalDescontos.Text := '0,00';
  edSalarioLiquido.Text := '0,00';
  wSalarioBase := 0;
  edCodigo.Clear;

  cbNome.Enabled := True;

  // cbNome.SetFocus;
end;

procedure TfrFolhaPagamento.pLimparCamposCadastroFuncionario;
begin
  edCodigoFuncionario.Clear;
  edNomeFuncionario.Clear;
  cbCargo.ItemIndex := -1;
  edEndereco.Clear;
  edTelefone.Clear;
end;

// Validacoes

function TfrFolhaPagamento.fValidarCamposFuncionario: Boolean;

begin
  Result := True;
  if wCodigoFuncionario = 0 then
     begin
       ShowMessage('Informe o código do funcionário');
       Result := False;
     end
  else if wNomeFuncionario = '' then
     begin
       ShowMessage('Informe o nome do funcionário');
       Result := False;
     end
  else if wCargoFuncionario = '' then
     begin
       ShowMessage('Informe o cargo do funcionário');
       Result := False;
     end
  else if wEnderecoFuncionario = '' then
     begin
       ShowMessage('Informe o endereço do funcionário');
       Result := False;
     end
  else if wTelefoneFuncionario = '' then
     begin
       ShowMessage('Informe o telefone do funcionário');
       Result := False;
     end;
end;

function TfrFolhaPagamento.fValidarCampoSalarioBase: Boolean;
begin
  Result := True;
  if wSalarioBase = 0 then
     begin
       ShowMessage('Informe o salário base');
       Result := False;
     end;
end;

function TfrFolhaPagamento.fValidarCamposFolha: Boolean;
begin
  wMesFolha    := cbMes.Text;
  wAnoFolha    := StrToIntDef(edAno.Text,0);
  wNomeFolha   := cbNome.Text;
  wCargoFolha  := edCargo.Text;

  Result := True;
  if wNomeFolha = '' then
     begin
       ShowMessage('Informe o nome');
       Result := False;
     end
  else if wCargoFolha = '' then
     begin
       ShowMessage('Informe o cargo');
       Result := False;
     end
  else if wMesFolha = '' then
     begin
       ShowMessage('Informe o mês');
       Result := False;
     end
  else if wAnoFolha = 0 then
     begin
       ShowMessage('Informe o ano');
       Result := False;
     end;
end;

// novas adicoes

function TfrFolhaPagamento.fCalcularInss: Currency;
var
  wPrimeiraFaixa: Currency;
  wSegundaFaixa: Currency;
  wTerceiraFaixa: Currency;
  wQuartaFaixa: Currency;
begin
  wSegundaFaixa:= 0;
  wTerceiraFaixa:= 0;
  wQuartaFaixa:= 0;

  if wSalarioBase <= 1621 then
     begin
      wPrimeiraFaixa := wSalarioBase * 0.075;
     end
  else if (wSalarioBase > 1621) and (wSalarioBase <= 2902.84) then
    begin
      wPrimeiraFaixa := fCalcularPrimeiraFaixa;
      wSegundaFaixa  := fCalcularSegundaFaixa;
     end
  else if (wSalarioBase > 2902.84) and (wSalarioBase <= 4354.27) then
     begin
      wPrimeiraFaixa := fCalcularPrimeiraFaixa;
      wSegundaFaixa  := fCalcularSegundaFaixa;
      wTerceiraFaixa := fCalcularTerceiraFaixa;
     end
  else if (wSalarioBase > 4354.27) and (wSalarioBase <= 8157.41) then
     begin
      wPrimeiraFaixa := fCalcularPrimeiraFaixa;
      wSegundaFaixa  := fCalcularSegundaFaixa;
      wTerceiraFaixa := fCalcularTerceiraFaixa;
      wQuartaFaixa   := (wSalarioBase - 4354.27) * 0.14;
     end
  else
     begin
       wPrimeiraFaixa := fCalcularPrimeiraFaixa;
       wSegundaFaixa  := fCalcularSegundaFaixa;
       wTerceiraFaixa := fCalcularTerceiraFaixa;
       wQuartaFaixa   := (8157.41 - 4354.27) * 0.14;
     end;

  Result := wPrimeiraFaixa + wSegundaFaixa + wTerceiraFaixa + wQuartaFaixa;
end;

function TfrFolhaPagamento.fCalcularPrimeiraFaixa: Currency;
begin
  Result := 1621 * 0.075;
end;

function TfrFolhaPagamento.fCalcularSegundaFaixa: Currency;
begin
  Result := (2902.84 - 1621) * 0.09;
end;

function TfrFolhaPagamento.fCalcularTerceiraFaixa: Currency;
begin
  Result := (4354.27 - 2902.84) * 0.12;
end;

procedure TfrFolhaPagamento.edCodigoEnter(Sender: TObject);
begin
  pLimparCampos;
end;

procedure TfrFolhaPagamento.pAbrirConsulta;
begin
  pnConsultaFuncionarios.Visible := True;
  pnConsultaFuncionarios.Top  := 100;
  pnConsultaFuncionarios.Left := 100;
end;

procedure TfrFolhaPagamento.btDeletarFuncionarioClick(Sender: TObject);
begin
  pDeletarFuncionario;
end;

procedure TfrFolhaPagamento.pDeletarFuncionario;
begin
  cdsFuncionario.IndexFieldNames := 'bdCODIGOFUNCIONARIO';

  if cdsFuncionario.FindKey([edCodigoFuncionario.Text]) then
     begin
       if MessageDlg('Deseja excluir este funcionário?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            cdsFuncionario.Delete;
            ShowMessage('Funcionário deletado com sucesso');
            pLimparCamposCadastroFuncionario;
            btDeletarFuncionario.Enabled := False;
          end
       else
         begin
           Exit;
         end;
     end
   else
     begin
        ShowMessage('Funcionário não encontrado');
     end
end;

procedure TfrFolhaPagamento.pDeletarFolha;
begin
  cdsFolhaPagamento.IndexFieldNames := 'bdCODIGO';

  if cdsFolhaPagamento.FindKey([edCodigo.Text]) then
     begin
        if MessageDlg('Deseja excluir esta folha?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
           begin
             cdsFolhaPagamento.Delete;
             ShowMessage('Folha deletada com sucesso');
             pLimparCampos;
             btDeletarFolha.Enabled := False;
           end
        else
           begin
             Exit;
           end;
      end
   else
      begin
        ShowMessage('Funcionário não encontrado');
      end
end;

procedure TfrFolhaPagamento.btDeletarFolhaClick(Sender: TObject);
begin
  pDeletarFolha;
end;

end.
