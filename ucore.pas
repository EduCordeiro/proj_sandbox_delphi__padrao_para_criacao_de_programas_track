unit ucore;

interface

uses
  Windows, Messages, Variants, Graphics, Controls, FileCtrl,
  Dialogs, StdCtrls,  Classes, SysUtils, Forms,
  DB, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZSqlProcessor,
  ADODb, DBTables,
  udatatypes_apps,
  // Classes
  ClassParametrosDeEntrada,
  ClassArquivoIni, ClassStrings, ClassConexoes, ClassConf, ClassMySqlBases,
  ClassTextFile, ClassDirectory, ClassLog, ClassFuncoesWin, ClassLayoutArquivo,
  ClassBlocaInteligente, ClassFuncoesBancarias, ClassPlanoDeTriagem, ClassExpressaoRegular,
  ClassStatusProcessamento, ClassDateTime, ClassSMTPDelphi;

type

  TCore = class(TObject)
  private

    __queryMySQL_processamento__    : TZQuery;
    __queryMySQL_processamento2__   : TZQuery;
    __queryMySQL_Insert_            : TZQuery;
    __queryMySQL_plano_de_triagem__ : TZQuery;

      procedure StoredProcedure_Dropar(Nome: string; logBD:boolean=false; idprograma:integer=0);

      function StoredProcedure_Criar(Nome : string; scriptSQL: TStringList): boolean;

      procedure StoredProcedure_Executar(Nome: string; ComParametro:boolean=false; logBD:boolean=false; idprograma:integer=0);

      function Compactar_Arquivo_7z(Arquivo, destino : String; mover_arquivo: Boolean=false): integer;
      function Extrair_Arquivo_7z(Arquivo, destino : String): integer;

      PROCEDURE COMPACTAR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String; MOVER_ARQUIVO: Boolean=FALSE);
      PROCEDURE EXTRAIR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String);

      procedure Atualiza_arquivo_conf_C(ArquivoConf, sINP, sOUT, sTMP, sLOG, sRGP: String);
      procedure execulta_app_c(app, arquivo_conf: string);

  public

    __ListaPlanoDeTriagem__       : TRecordPlanoTriagemCorreios;

    objParametrosDeEntrada   : TParametrosDeEntrada;
    objConexao               : TMysqlDatabase;
    objPlanoDeTriagem        : TPlanoDeTriagem;
    objString                : TFormataString;
    objLogar                 : TArquivoDelog;
    objDateTime              : TFormataDateTime;
    objArquivoIni            : TArquivoIni;
    objArquivoDeConexoes     : TArquivoDeConexoes;
    objArquivoDeConfiguracao : TArquivoConf;
    objDiretorio             : TDiretorio;
    objFuncoesWin            : TFuncoesWin;
    objLayoutArquivoCliente  : TLayoutCliente;
    objBlocagemInteligente   : TBlocaInteligente;
    objFuncoesBancarias      : TFuncoesBancarias;
    objExpressaoRegular      : TExpressaoRegular;
    objStatusProcessamento   : TStausProcessamento;
    objEmail                 : TSMTPDelphi;

  // FUNÇÃO DE PROCESSAMENTO
    Procedure PROCESSAMENTO();

    PROCEDURE COMPACTAR();
    PROCEDURE EXTRAIR();

    function GERA_LOTE_PEDIDO(): String;
    Procedure VALIDA_LOTE_PEDIDO();
    Procedure AtualizaDadosTabelaLOG();

    function PesquisarLote(LOTE_PEDIDO : STRING; status : Integer): Boolean;

    procedure ExcluirBase(NomeTabela: String);
    procedure ExcluirTabela(NomeTabela: String);
    function EnviarEmail(Assunto: string=''; Corpo: string=''): Boolean;
    procedure MainLoop();
    constructor create();

    procedure ReverterArquivos();

    procedure getListaDeArquivosPendentes();
    procedure getListaDeArquivosTrack();
    procedure getListaDeArquivosJaProcessados();

    function ArquivoExieteTabelaTrackLine(Arquivo: string): Boolean;
    procedure CriaMovimento();

  end;

implementation

uses uMain, Math;

constructor TCore.create();
var
  sMSG                       : string;
  sArquivosScriptSQL         : string;
  stlScripSQL                : TStringList;
begin

  try

    stlScripSQL                                              := TStringList.Create();

    objStatusProcessamento                                   := TStausProcessamento.create();
    objParametrosDeEntrada                                   := TParametrosDeEntrada.Create();

    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_PENDENTES      := TStringList.Create();
    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_TRACK          := TStringList.Create();
    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS := TStringList.Create();
    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER       := TStringList.Create();

    objLogar                                                 := TArquivoDelog.Create();
    if FileExists(objLogar.getArquivoDeLog()) then
      objFuncoesWin.DelFile(objLogar.getArquivoDeLog());

    objFuncoesWin                        := TFuncoesWin.create(objLogar);
    objString                            := TFormataString.Create(objLogar);
    objDateTime                          := TFormataDateTime.Create(objLogar);
    objLayoutArquivoCliente              := TLayoutCliente.Create();
    objFuncoesBancarias                  := TFuncoesBancarias.Create();
    objExpressaoRegular                  := TExpressaoRegular.Create();

    objArquivoIni                        := TArquivoIni.create(objLogar,
                                                               objString,
                                                               ExtractFilePath(Application.ExeName),
                                                               ExtractFileName(Application.ExeName));

    objArquivoDeConexoes                 := TArquivoDeConexoes.create(objLogar,
                                                                      objString,
                                                                      objArquivoIni.getPathConexoes());

    objArquivoDeConfiguracao             := TArquivoConf.create(objArquivoIni.getPathConfiguracoes(),
                                                                ExtractFileName(Application.ExeName));

    objParametrosDeEntrada.ID_PROCESSAMENTO := objArquivoDeConfiguracao.getIDProcessamento;

    objConexao                           := TMysqlDatabase.Create();

    if objArquivoIni.getPathConfiguracoes() <> '' then
    begin

      objParametrosDeEntrada.PATHENTRADA                                := objArquivoDeConfiguracao.getConfiguracao('path_default_arquivos_entrada');
      objParametrosDeEntrada.PATHSAIDA                                  := objArquivoDeConfiguracao.getConfiguracao('path_default_arquivos_saida');
      objParametrosDeEntrada.TABELA_PROCESSAMENTO                       := objArquivoDeConfiguracao.getConfiguracao('TABELA_PROCESSAMENTO');
      objParametrosDeEntrada.TABELA_PROCESSAMENTO2                      := objArquivoDeConfiguracao.getConfiguracao('TABELA_PROCESSAMENTO2');
      objParametrosDeEntrada.TABELA_LOTES_PEDIDOS                       := objArquivoDeConfiguracao.getConfiguracao('TABELA_LOTES_PEDIDOS');
      objParametrosDeEntrada.TABELA_PLANO_DE_TRIAGEM                    := objArquivoDeConfiguracao.getConfiguracao('tabela_plano_de_triagem');
      objParametrosDeEntrada.CARREGAR_PLANO_DE_TRIAGEM_MEMORIA          := objArquivoDeConfiguracao.getConfiguracao('CARREGAR_PLANO_DE_TRIAGEM_MEMORIA');
      objParametrosDeEntrada.TABELA_BLOCAGEM_INTELIGENTE                := objArquivoDeConfiguracao.getConfiguracao('TABELA_BLOCAGEM_INTELIGENTE');
      objParametrosDeEntrada.TABELA_BLOCAGEM_INTELIGENTE_RELATORIO      := objArquivoDeConfiguracao.getConfiguracao('TABELA_BLOCAGEM_INTELIGENTE_RELATORIO');
      objParametrosDeEntrada.TABELA_ENTRADA_SP                          := objArquivoDeConfiguracao.getConfiguracao('TABELA_ENTRADA_SP');
      objParametrosDeEntrada.TABELA_AUX_SP                              := objArquivoDeConfiguracao.getConfiguracao('TABELA_AUX_SP');
      objParametrosDeEntrada.LIMITE_DE_SELECT_POR_INTERACOES_NA_MEMORIA := objArquivoDeConfiguracao.getConfiguracao('numero_de_select_por_interacoes_na_memoria');
      objParametrosDeEntrada.NUMERO_DE_IMAGENS_PARA_BLOCAGENS           := objArquivoDeConfiguracao.getConfiguracao('NUMERO_DE_IMAGENS_PARA_BLOCAGENS');
      objParametrosDeEntrada.BLOCAR_ARQUIVO                             := objArquivoDeConfiguracao.getConfiguracao('BLOCAR_ARQUIVO');
      objParametrosDeEntrada.BLOCAGEM                                   := objArquivoDeConfiguracao.getConfiguracao('BLOCAGEM');
      objParametrosDeEntrada.MANTER_ARQUIVO_ORIGINAL                    := objArquivoDeConfiguracao.getConfiguracao('MANTER_ARQUIVO_ORIGINAL');
      objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO                     := objArquivoDeConfiguracao.getConfiguracao('FORMATACAO_LOTE_PEDIDO');
      objParametrosDeEntrada.lista_de_caracteres_invalidos              := objArquivoDeConfiguracao.getConfiguracao('lista_de_caracteres_invalidos');
      objParametrosDeEntrada.SP_001                                     := objArquivoDeConfiguracao.getConfiguracao('SP_001');
      objParametrosDeEntrada.SP_002                                     := objArquivoDeConfiguracao.getConfiguracao('SP_002');
      objParametrosDeEntrada.SP_003                                     := objArquivoDeConfiguracao.getConfiguracao('SP_003');
      objParametrosDeEntrada.SP_001_NAME                                := objArquivoDeConfiguracao.getConfiguracao('SP_001_NAME');
      objParametrosDeEntrada.SP_002_NAME                                := objArquivoDeConfiguracao.getConfiguracao('SP_002_NAME');
      objParametrosDeEntrada.SP_003_NAME                                := objArquivoDeConfiguracao.getConfiguracao('SP_003_NAME');
      objParametrosDeEntrada.eHost                                      := objArquivoDeConfiguracao.getConfiguracao('eHost');
      objParametrosDeEntrada.eUser                                      := objArquivoDeConfiguracao.getConfiguracao('eUser');
      objParametrosDeEntrada.eFrom                                      := objArquivoDeConfiguracao.getConfiguracao('eFrom');
      objParametrosDeEntrada.eTo                                        := objArquivoDeConfiguracao.getConfiguracao('eTo');

      objParametrosDeEntrada.EXTENCAO_ARQUIVOS                          := objArquivoDeConfiguracao.getConfiguracao('EXTENCAO_ARQUIVOS');

      objParametrosDeEntrada.OF_FORMULARIO                              := objArquivoDeConfiguracao.getConfiguracao('OF_FORMULARIO');
      objParametrosDeEntrada.PESO_PAPEL                                 := objArquivoDeConfiguracao.getConfiguracao('PESO_PAPEL');
      objParametrosDeEntrada.ACABAMENTO                                 := objArquivoDeConfiguracao.getConfiguracao('ACABAMENTO');
      objParametrosDeEntrada.PAPEL                                      := objArquivoDeConfiguracao.getConfiguracao('PAPEL');

      objParametrosDeEntrada.CRIAR_CSV_TRACK                            := StrTobool(objArquivoDeConfiguracao.getConfiguracao('CRIAR_CSV_TRACK'));
      objParametrosDeEntrada.PATH_TRACK_FATURAMENTO                     := objArquivoDeConfiguracao.getConfiguracao('PATH_TRACK_FATURAMENTO');

      objParametrosDeEntrada.TABELA_TRACK                               := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK');
      objParametrosDeEntrada.TABELA_TRACK_LINE                          := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK_LINE');
      objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY                  := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK_LINE_HISTORY');

      objParametrosDeEntrada.APP_C_GERA_SPOOL_EXE                         := objArquivoDeConfiguracao.getConfiguracao('APP_C_GERA_SPOOL_EXE');
      objParametrosDeEntrada.APP_C_GERA_SPOOL_CFG                         := objArquivoDeConfiguracao.getConfiguracao('APP_C_GERA_SPOOL_CFG');

      objParametrosDeEntrada.app_7z_32bits                              := objArquivoDeConfiguracao.getConfiguracao('app_7z_32bits');
      objParametrosDeEntrada.app_7z_64bits                              := objArquivoDeConfiguracao.getConfiguracao('app_7z_64bits');
      objParametrosDeEntrada.ARQUITETURA_WINDOWS                        := objArquivoDeConfiguracao.getConfiguracao('ARQUITETURA_WINDOWS');

      objParametrosDeEntrada.LOGAR                                      := objArquivoDeConfiguracao.getConfiguracao('LOGAR');

      //================
      //  LOGA USUÁRIO
      //========================================================================================================================================================
      objParametrosDeEntrada.APP_LOGAR                                  := objArquivoDeConfiguracao.getConfiguracao('APP_LOGAR');
      objParametrosDeEntrada.TABELA_LOTES_PEDIDOS_LOGIN                 := objArquivoDeConfiguracao.getConfiguracao('TABELA_LOTES_PEDIDOS_LOGIN');
      //========================================================================================================================================================

      objParametrosDeEntrada.ENVIAR_EMAIL                               := objArquivoDeConfiguracao.getConfiguracao('ENVIAR_EMAIL');



      objLogar.Logar('[DEBUG] TfrmMain.FormCreate() - Versão do programa: ' + objFuncoesWin.GetVersaoDaAplicacao());

      objParametrosDeEntrada.PathArquivo_TMP := objArquivoIni.getPathArquivosTemporarios();

      // Criando a Conexao
      objConexao.ConectarAoBanco(objArquivoDeConexoes.getHostName,
                                 'mysql',
                                 objArquivoDeConexoes.getUser,
                                 objArquivoDeConexoes.getPassword,
                                 objArquivoDeConexoes.getProtocolo
                                 );

      sArquivosScriptSQL := ExtractFileName(Application.ExeName);
      sArquivosScriptSQL := StringReplace(sArquivosScriptSQL, '.exe', '.sql', [rfReplaceAll, rfIgnoreCase]);

      stlScripSQL.LoadFromFile(objArquivoIni.getPathScripSQL() + sArquivosScriptSQL);
      objConexao.ExecutaScript(stlScripSQL);

      objBlocagemInteligente   := TBlocaInteligente.create(objParametrosDeEntrada,
                                                           objConexao,
                                                           objFuncoesWin,
                                                           objString,
                                                           objLogar);

      // Criando Objeto de Plano de Triagem
      if StrToBool(objParametrosDeEntrada.CARREGAR_PLANO_DE_TRIAGEM_MEMORIA) then
        objPlanoDeTriagem := TPlanoDeTriagem.create(objConexao,
                                                    objLogar,
                                                    objString,
                                                    objParametrosDeEntrada.TABELA_PLANO_DE_TRIAGEM, fac);



      objParametrosDeEntrada.stlRelatorioQTDE           := TStringList.Create();

      // LISTA DE ARUQIVOS JA PROCESSADOS
      getListaDeArquivosJaProcessados();


      objParametrosDeEntrada.STL_LOG_TXT                := TStringList.Create(); 

      IF StrToBool(objParametrosDeEntrada.LOGAR) THEN
      BEGIN

          //================
          //  LOGA USUÁRIO
          //==========================================================================================================================================================
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_TAB_INDEX      := '2';
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_NOME_APLICACAO := StringReplace(ExtractFileName(Application.ExeName), '.EXE', '', [rfReplaceAll, rfIgnoreCase]);
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR  := ExtractFilePath(Application.ExeName) +
                                                                       StringReplace(ExtractFileName(objParametrosDeEntrada.APP_LOGAR), '.EXE', '.TXT', [rfReplaceAll, rfIgnoreCase]);

          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR  := StringReplace(objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR, '\', '/', [rfReplaceAll, rfIgnoreCase]);

          

          objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO := TStringList.Create();
          objFuncoesWin.ExecutarPrograma(objParametrosDeEntrada.APP_LOGAR
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_TAB_INDEX
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_NOME_APLICACAO
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR);

          objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.LoadFromFile(objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR);

          //=====================
          //   CAMPOS DE LOGIN
          //=====================
          objParametrosDeEntrada.USUARIO_LOGADO_APP           := objString.getTermo(1, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          := objString.getTermo(2, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_LOTE               := objString.getTermo(3, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN := objString.getTermo(4, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_IP                 := objString.getTermo(5, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_ID                 := objString.getTermo(6, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);

          IF (Trim(objParametrosDeEntrada.USUARIO_LOGADO_APP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_CHAVE_APP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_LOTE) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_IP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_ID) ='')
          THEN
            objParametrosDeEntrada.USUARIO_LOGADO_APP := '-1';
      END;

      //=========================
      //    DADOS DE REDE APP
      //=========================
      objParametrosDeEntrada.HOSTNAME                     := objFuncoesWin.getNetHostName;
      objParametrosDeEntrada.IP                           := objFuncoesWin.GetIP;
      objParametrosDeEntrada.USUARIO_SO                   := objFuncoesWin.GetUsuarioLogado;

      //========================
      //  GERA LOTE PEDIDO
      //========================
      if NOT StrToBool(objParametrosDeEntrada.LOGAR) then
      BEGIN

        objParametrosDeEntrada.PEDIDO_LOTE                  := GERA_LOTE_PEDIDO();

        objParametrosDeEntrada.USUARIO_LOGADO_APP           := objParametrosDeEntrada.USUARIO_SO;
        objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          := objParametrosDeEntrada.ID_PROCESSAMENTO;
        objParametrosDeEntrada.APP_LOGAR_LOTE               := objParametrosDeEntrada.PEDIDO_LOTE;
        objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN := objParametrosDeEntrada.USUARIO_SO;
        objParametrosDeEntrada.APP_LOGAR_IP                 := objParametrosDeEntrada.IP;
        objParametrosDeEntrada.APP_LOGAR_ID                 := objParametrosDeEntrada.ID_PROCESSAMENTO;

      END
      ELSE
      IF objParametrosDeEntrada.USUARIO_LOGADO_APP <> '-1' THEN
        objParametrosDeEntrada.PEDIDO_LOTE                := GERA_LOTE_PEDIDO();
      //==========================================================================================================================================================

    end;

  except
    on E:Exception do
    begin

      sMSG := '[ERRO] Não foi possível inicializar as configurações aq do programa. '+#13#10#13#10
            + ' EXCEÇÃO: '+E.Message+#13#10#13#10
            + ' O programa será encerrado agora.';

      showmessage(sMSG);

      objLogar.Logar(sMSG);

      Application.Terminate;
    end;
  end;

end;

function TCore.GERA_LOTE_PEDIDO(): String;
var
  sComando : string;
  sData    : string;
begin

  //==================
  //  CRIA NOVO LOTE
  //==================
  sData := FormatDateTime('YYYY-MM-DD hh:mm:ss', Now());

  sComando := ' insert into ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS + '(VALIDO, DATA_CRIACAO, CHAVE, USUARIO_WIN, USUARIO_APP, IP, ID, LOTE_LOGIN, HOSTNAME)'
            + ' Value('
                      + '"'   + 'N'
                      + '","' + sData
                      + '","' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP
                      + '","' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN
                      + '","' + objParametrosDeEntrada.USUARIO_LOGADO_APP
                      + '","' + objParametrosDeEntrada.APP_LOGAR_IP
                      + '","' + objParametrosDeEntrada.ID_PROCESSAMENTO
                      + '","' + objParametrosDeEntrada.APP_LOGAR_LOTE
                      + '","' + objParametrosDeEntrada.HOSTNAME
                      + '")';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  //========================
  //  RETORNA LOTE CRIADO
  //========================
  sComando := ' SELECT LOTE_PEDIDO FROM  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' WHERE '
                      + '     VALIDO        = "' + 'N'                                                 + '"'
                      + ' AND DATA_CRIACAO  = "' + sData                                               + '"'
                      + ' AND CHAVE         = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          + '"'
                      + ' AND USUARIO_WIN   = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN + '"'
                      + ' AND USUARIO_APP   = "' + objParametrosDeEntrada.USUARIO_LOGADO_APP           + '"'
                      + ' AND HOSTNAME      = "' + objParametrosDeEntrada.HOSTNAME                     + '"'
                      + ' AND LOTE_LOGIN    = "' + objParametrosDeEntrada.APP_LOGAR_LOTE               + '"'
                      + ' AND IP            = "' + objParametrosDeEntrada.APP_LOGAR_IP                 + '"'
                      + ' AND ID            = "' + objParametrosDeEntrada.ID_PROCESSAMENTO             + '"';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  Result := FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, __queryMySQL_processamento__.FieldByName('LOTE_PEDIDO').AsInteger);

end;

PROCEDURE TCore.VALIDA_LOTE_PEDIDO();
VAR
  sComando                : string;
BEGIN

  //========================
  //  RETORNA LOTE CRIADO
  //========================
  sComando := ' UPDATE  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' set VALIDO         = "' + objParametrosDeEntrada.STATUS_PROCESSAMENTO  + '"'
            + '    ,RELATORIO_QTD  = "' + objParametrosDeEntrada.stlRelatorioQTDE.Text + '"'
            + '    ,LOTE_LOGIN     = "' + objParametrosDeEntrada.APP_LOGAR_LOTE    + '"'
            + ' WHERE '
            + '     LOTE_PEDIDO   = "' + objParametrosDeEntrada.PEDIDO_LOTE                   + '"'
            + ' AND VALIDO        = "' + 'N'                                                  + '"'
            + ' AND CHAVE         = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP           + '"'
            + ' AND USUARIO_WIN   = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN  + '"'
            + ' AND HOSTNAME      = "' + objParametrosDeEntrada.HOSTNAME                      + '"'
            + ' AND IP            = "' + objParametrosDeEntrada.APP_LOGAR_IP                  + '"'
            + ' AND ID            = "' + objParametrosDeEntrada.ID_PROCESSAMENTO              + '"';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

end;

Procedure TCore.AtualizaDadosTabelaLOG();
var
  sComando                  : String;
Begin
  //=========================================================================
  //  GRAVA LOG NA TABELA DE LOGIN - SOMENTE SE O PARÂMETRO LOGAR FOR TRUE
  //=========================================================================
  if StrToBool(objParametrosDeEntrada.LOGAR) then
  begin
    objParametrosDeEntrada.STL_LOG_TXT.Text := StringReplace(objParametrosDeEntrada.STL_LOG_TXT.Text, '\', '\\', [rfReplaceAll, rfIgnoreCase]);

    sComando := ' update ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS_LOGIN
              + ' SET '
              + '      LOG_APP          = "' + objParametrosDeEntrada.STL_LOG_TXT.Text                           + '"'
              + '     ,VALIDO           = "' + objParametrosDeEntrada.STATUS_PROCESSAMENTO                       + '"'
              + '     ,QTD_PROCESSADA   = "' + IntToStr(objParametrosDeEntrada.TOTAL_PROCESSADOS_LOG)            + '"'
              + '     ,QTD_INVALIDOS    = "' + IntToStr(objParametrosDeEntrada.TOTAL_PROCESSADOS_INVALIDOS_LOG)  + '"'
              + '     ,LOTE_APP         = "' + objParametrosDeEntrada.PEDIDO_LOTE                                + '"'
              + '     ,RELATORIO_QTD    = "' + objParametrosDeEntrada.stlRelatorioQTDE.Text                      + '"'
              + ' WHERE CHAVE       = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          + '"'
              + '   AND LOTE_PEDIDO = "' + objParametrosDeEntrada.APP_LOGAR_LOTE               + '"'
              + '   AND USUARIO_WIN = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN + '"'
              + '   AND USUARIO_APP = "' + objParametrosDeEntrada.USUARIO_LOGADO_APP           + '"'
              + '   AND HOSTNAME    = "' + objParametrosDeEntrada.HOSTNAME                     + '"'
              + '   AND IP          = "' + objParametrosDeEntrada.APP_LOGAR_IP                 + '"'
              + '   AND ID          = "' + objParametrosDeEntrada.APP_LOGAR_ID                 + '"';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
  end;

end;

procedure TCore.MainLoop();
var
  sMSG : string;
begin

  objLogar.Logar('[DEBUG] TCore.MainLoop() - begin...');
  try
    try

        objDiretorio := TDiretorio.create(objParametrosDeEntrada.PathEntrada);
        objParametrosDeEntrada.PathEntrada := objDiretorio.getDiretorio();

        objDiretorio.setDiretorio(objParametrosDeEntrada.PathSaida);
        objParametrosDeEntrada.PathSaida   := objDiretorio.getDiretorio();

//      PROCESSAMENTO();

    finally

      if Assigned(objDiretorio) then
      begin
        objDiretorio.destroy;
        Pointer(objDiretorio) := nil;
      end;

    end;

  except

    // 0------------------------------------------0
    // |  Excessões desntro do objCore caem aqui  |
    // 0------------------------------------------0
    on E:Exception do
    begin

      sMSG :='Erro ao execultar a Função MainLoop(). ' + #13#10#13#10
                 +'EXCEÇÃO: '+E.Message+#13#10#13#10
                 +'O programa será encerrado agora.';

      IF StrToBool(objParametrosDeEntrada.ENVIAR_EMAIL) THEN
        EnviarEmail('ERRO DE PROCESSAMENTO !!!', sMSG + #13 + #13 + 'SEGUE LOG EM ANEXO.' + #13 + #13
        + 'DETALHES DE LOGIN' + #13
        + '=================' + #13
        + 'HOSTNAME.......................: ' + objParametrosDeEntrada.HOSTNAME + #13
        + 'USUARIO LOGADO.................: ' + objParametrosDeEntrada.USUARIO_LOGADO_APP + #13
        + 'USUARIO SO.....................: ' + objParametrosDeEntrada.USUARIO_SO + #13
        + 'LOTE LOGIN.....................: ' + objParametrosDeEntrada.APP_LOGAR_LOTE + #13
        + 'IP.............................: ' + objParametrosDeEntrada.IP);

      showmessage(sMSG);
      objLogar.Logar(sMSG);

    end;
  end;

  objLogar.Logar('[DEBUG] TCore.MainLoop() - ...end');

end;

Procedure TCore.PROCESSAMENTO();
Var

Arq_Arquivo_Entada               : TextFile;

sArquivoEntrada                  : string;

sPathCsvTrackFaturamento         : string;
sPathCsvTrackBKP                 : string;

sValues                          : string;
sComando                         : string;
sCampos                          : string;
sCampoEncarte                    : string;
sLinha                           : string;

sMovimento                       : string;
//sListaDeOfEnvelopePorOf          : string;
sListaDeOfEncartePorOf           : string;
sListaDePortesPorOf              : string;
//sOfFormulario                    : string;
//sOfEncartes                      : string;
//sOfEnvelope                      : string;
sPorte                           : string;
sLogo                            : string;
sQTD                             : string;
sGrupoFolhas                     : string;
sTotalGrupoFolhas                : string;
sLote                            : string;
sDataPostagem                    : string;

sTrackMovimentoRecebimento       : string;
sTrackMovimentoProcessamento     : string;
sTrackEmpresa                    : string;
sTrackPacote                     : string;
sTrackCodigoI                    : string;
sTrackDescricao                  : string;
sTrackCodigoMatriz               : string;
sTrackGestor                     : string;
sTrackCartao                     : string;
sTrackPapel                      : string;
sTrackAcabamento                 : string;
sTrackTransmissao                : string;
sTrackObjetos                    : string;
sTrackPaginas                    : string;
sTrackFolhas                     : string;
sTrackOF_Formulario              : string;
sTrackOF_Envelope                : string;
sTrackDataLoteQtdPostagem        : string;
sTrackTarefa                     : string;
sTrackTipoArquivo                : string;
sTrackCepInvalido                : string;

//sTrackPlataforma                 : string;
//sTrackDataCorte                  : string;
//sTrackDataVencimento             : string;

//sTrackProduto                    : string;
//sTrackArquivo                    : string;
//sTrackDescricao                  : string;
//sTrackDescricaoSubProduto        : string;
//sTrackTransmissaoInicio          : string;
//sTrackTransmissaoFin             : string;
//sTrackProcessamentoInicio        : string;
//sTrackProcessamentoFin           : string;
//sTrackImpressaoInicio            : string;
//sTrackImpressaoFin               : string;
//sTrackAcabamentoInicio           : string;
//sTrackAcabamentoFin              : string;
//sTrackTipoPapel                  : string;
//sTrackTipoImpressao              : string;
//sTrackDuplex                     : string;

//sTrackQuantidadeImagens          : string;

//sTrackDataCif                    : string;

sTrackPortes                     : string;

sTrackLotes                      : string;


//iQuantidadeDeObjetosPorOf        : Integer;
//iQuantidadeDeFolhasPorOf         : Integer;

iTotalLocal                      : Integer;
iTotalEstadual                   : Integer;
iTotalNacional                   : Integer;

iTotalObjetosRecebidosNoArquivo  : Integer;

sCabecalho                       : string;
sTimeStamp                       : string;

//sV1                              : string;

iDirecao                         : Integer;

ListaOfPorArquivo                : TStringList;

iQTD                             : Integer;

iContArquivos                    : Integer;
iTotalDeArquivos                 : Integer;
iContSequenciaEncarte            : Integer;

// Variáveis de controle do select
iTotalDeRegistrosDaTabela   : Integer;
iLimit                      : Integer;
iTotalDeInteracoesDeSelects : Integer;
iResto                      : Integer;
iRegInicial                 : Integer;
iQtdeRegistros              : Integer;
iContInteracoesDeSelects    : Integer;

sOperadora                  : string;
sContrato                   : string;
sCep                        : string;

begin

  objParametrosDeEntrada.TIMESTAMP := Now();
  sTimeStamp                       := FormatDateTime('YYYYMMDD', objParametrosDeEntrada.TimeStamp);
  sPathCsvTrackFaturamento         := objString.AjustaPath(objParametrosDeEntrada.PATH_TRACK_FATURAMENTO);

  sCabecalho :=   'OF_FORMULARIO'
               + ';OF_ENVELOPE'
               + ';OF_ENCATE'
               + ';MOVIMENTO'
               + ';FILLER'
               + ';ARQUIVO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';TIMESTAMP'
               + ';ACABAMENTO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';LOTE_PROCESAMENTO'
               + ';QUANTIDADE_DE_OBJETOS_POR_OF'
               + ';QUANTIDADE_DE_FOLHAS_POR_OF'
               + ';QUANTIDADE_DE_PAGINAS_POR_OF'
               + ';FILLER'
               + ';FILLER'
               + ';CARTAO_POSTAGEM'
               + ';DATA_LOTE_QTD_POSTAGEM'
               + ';TOTAL_LOCAL'
               + ';TOTAL_ESTADUAL'
               + ';TOTAL_NACIONAL'
               + ';TOTAL'
               + ';PORTE[GR]'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';PAPEL'
               ;
  ListaOfPorArquivo := TStringList.Create();

  objParametrosDeEntrada.stlRelatorioQTDE.Clear;
  objParametrosDeEntrada.stlRelatorioQTDE.Add(sLinha);


  sComando := 'delete from ' + objParametrosDeEntrada.tabela_processamento;
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    ListaOfPorArquivo.Clear;
    ListaOfPorArquivo.Add(sCabecalho);

    sArquivoEntrada := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Strings[iContArquivos];

    //====================================================================================================================================================
    //  PEGA DADOS DE RECEBIMENTO
    //====================================================================================================================================================
    sComando := ' SELECT LOTE, MOVIMENTO '
              + ' FROM ' + objParametrosDeEntrada.TABELA_TRACK
              + ' WHERE ARQUIVO_AFP = "' + sArquivoEntrada + '"'
              + ' GROUP BY ARQUIVO_AFP ';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

    sTrackTarefa               := __queryMySQL_processamento__.FieldByName('LOTE').AsString;
    sTrackMovimentoRecebimento := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString;
    //====================================================================================================================================================

    //=====================================
    //  QUANTIDADES GERAIS POR OF
    //=====================================
    sComando := ' SELECT TIMESTAMP, MOVIMENTO, ACABAMENTO, PAPEL, OF_FORMULARIO, OF_ENVELOPE, COUNT(OF_FORMULARIO) AS OBJETOS, SUM(FOLHAS) AS FOLHAS, SUM(PAGINAS) AS PAGINAS '
              + ' FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
              + ' WHERE ARQUIVO_AFP = "' + sArquivoEntrada + '"'
              + ' GROUP BY MOVIMENTO, ACABAMENTO, PAPEL, OF_FORMULARIO, OF_ENVELOPE ';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

    WHILE not __queryMySQL_processamento__.Eof DO
    BEGIN

      //=========================================================================================================
      // ATUALIZA O MOVIMENTO
      //=========================================================================================================
      sTrackMovimentoProcessamento      := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString;
      //=========================================================================================================

      sTrackTransmissao                 := __queryMySQL_processamento__.FieldByName('TIMESTAMP').AsString;
      sTrackAcabamento                  := __queryMySQL_processamento__.FieldByName('ACABAMENTO').AsString;
      sTrackPapel                       := __queryMySQL_processamento__.FieldByName('PAPEL').AsString;
      sTrackOF_Formulario               := __queryMySQL_processamento__.FieldByName('OF_FORMULARIO').AsString;
      sTrackOF_Envelope                 := __queryMySQL_processamento__.FieldByName('OF_ENVELOPE').AsString;
      sTrackObjetos                     := __queryMySQL_processamento__.FieldByName('OBJETOS').AsString;
      sTrackPaginas                     := __queryMySQL_processamento__.FieldByName('PAGINAS').AsString;
      sTrackFolhas                      := __queryMySQL_processamento__.FieldByName('FOLHAS').AsString;

      //====================================================================================================================================
      //   QUANTIDADE DE DIRECOES POR OF
      //====================================================================================================================================
      sPathCsvTrackBKP                  := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA)
                                        + sTrackMovimentoRecebimento + PathDelim
                                        + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(sTrackTarefa)) + PathDelim
                                        + 'TRACK' + PathDelim
                                        + 'PROCESSADO' + PathDelim;
      ForceDirectories(sPathCsvTrackBKP);
      //====================================================================================================================================

            //=====================================
            //   QUANTIDADE DE DIRECOES POR OF
            //=====================================
            sComando := ' SELECT DIRECAO, COUNT(DIRECAO) QTD FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                      + ' WHERE '
                      + '       ARQUIVO_AFP    = "' + sArquivoEntrada              + '"'
                      + '   AND ACABAMENTO     = "' + sTrackAcabamento             + '"'
                      + '   AND PAPEL          = "' + sTrackPapel                  + '"'
                      + '   AND OF_FORMULARIO  = "' + sTrackOF_Formulario          + '"'
                      + '   AND MOVIMENTO      = "' + sTrackMovimentoProcessamento + '"'
                      + '   AND DIRECAO <> "" '
                      + ' GROUP BY DIRECAO';
            objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

            iTotalLocal               := 0;
            iTotalEstadual            := 0;
            iTotalNacional            := 0;

            IF __queryMySQL_processamento2__.RecordCount > 0 THEN
            BEGIN

             // iQuantidadeDeObjetosPorOf := 0;

              while not __queryMySQL_processamento2__.Eof do
              begin

                iDirecao        := __queryMySQL_processamento2__.FieldByName('DIRECAO').AsInteger;
                iQTD            := __queryMySQL_processamento2__.FieldByName('QTD').AsInteger;

                case iDirecao of

                  1: iTotalLocal     := iTotalLocal    + iQTD;
                  2: iTotalEstadual  := iTotalEstadual + iQTD;
                  3: iTotalNacional  := iTotalNacional + iQTD;

                end;

               // iQuantidadeDeObjetosPorOf := iQuantidadeDeObjetosPorOf + iQTD;

                __queryMySQL_processamento2__.Next;

              END;

            end;

            //=====================================
            //   QUANTIDADE DE PORTES POR OF
            //=====================================
            sComando := ' SELECT PORTE, COUNT(PORTE) QTD FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                      + ' WHERE '
                      + '       ARQUIVO_AFP    = "' + sArquivoEntrada              + '"'
                      + '   AND ACABAMENTO     = "' + sTrackAcabamento             + '"'
                      + '   AND PAPEL          = "' + sTrackPapel                  + '"'
                      + '   AND OF_FORMULARIO  = "' + sTrackOF_Formulario          + '"'
                      + '   AND MOVIMENTO      = "' + sTrackMovimentoProcessamento + '"'
                      + '   AND PORTE         <> "" '
                      + ' group by PORTE';
            objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

            sTrackPortes := '';
            IF __queryMySQL_processamento2__.RecordCount > 0 THEN
            BEGIN

              while not __queryMySQL_processamento2__.Eof do
              begin

                sPorte   := __queryMySQL_processamento2__.FieldByName('PORTE').AsString;
                sQTD     := __queryMySQL_processamento2__.FieldByName('QTD').AsString;

                sTrackPortes := sTrackPortes + sPorte + ',' + sQTD;

                __queryMySQL_processamento2__.Next;

                if not __queryMySQL_processamento2__.Eof THEN
                  sTrackPortes := sTrackPortes + '|'

              END;

            end;

            //=====================================
            //   QUANTIDADE POR LOTES POR OF
            //=====================================
            sComando := ' SELECT CONCAT(MID(DATA_POSTAGEM, 1,2), "/", MID(DATA_POSTAGEM, 3,2), "/", MID(DATA_POSTAGEM, 5,2)) AS DATA_POSTAGEM, LOTE, COUNT(LOTE) QTD FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                      + ' WHERE '
                      + '       ARQUIVO_AFP    = "' + sArquivoEntrada              + '"'
                      + '   AND ACABAMENTO     = "' + sTrackAcabamento             + '"'
                      + '   AND PAPEL          = "' + sTrackPapel                  + '"'
                      + '   AND OF_FORMULARIO  = "' + sTrackOF_Formulario          + '"'
                      + '   AND MOVIMENTO      = "' + sTrackMovimentoProcessamento + '"'
                      + '   AND LOTE <> "" '
                      + ' group by LOTE';
            objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

            sTrackDataLoteQtdPostagem := '';
            IF __queryMySQL_processamento2__.RecordCount > 0 THEN
            BEGIN

              while not __queryMySQL_processamento2__.Eof do
              begin

                sDataPostagem := __queryMySQL_processamento2__.FieldByName('DATA_POSTAGEM').AsString;
                sLote         := __queryMySQL_processamento2__.FieldByName('LOTE').AsString;
                sQTD          := __queryMySQL_processamento2__.FieldByName('QTD').AsString;

                sTrackDataLoteQtdPostagem := sTrackDataLoteQtdPostagem + sDataPostagem + ',' + sLote + ',' + sQTD;

                __queryMySQL_processamento2__.Next;

                if not __queryMySQL_processamento2__.Eof THEN
                  sTrackDataLoteQtdPostagem := sTrackDataLoteQtdPostagem + '|'

              END;

            end;


      sLinha  :=       sTrackOF_Formulario                                                     // 'OF_FORMULARIO'
               + ';' + sTrackOF_Envelope                                                       // 'ListaDeOfEnvelopePorOf'
               + ';'                                                                           // 'ListaDeOfEncartePorOf'
               + ';' + sTrackMovimentoProcessamento                                            // 'MOVIMENTO'
               + ';'                                                                           // 'FILLER'
               + ';' + sArquivoEntrada                                                         // 'ARQUIVO'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';' + FormatDateTime('YYYY-MM-DD hh:mm:ss', objParametrosDeEntrada.TIMESTAMP) // 'TIMESTAMP'
               + ';' + sTrackAcabamento                                                        // 'ACABAMENTO'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';' + sTrackTarefa                                                            // 'LOTE_PROCESAMENTO'
               + ';' + sTrackObjetos                                                           // 'QUANTIDADE_DE_OBJETOS_POR_OF'
               + ';' + sTrackFolhas                                                            // 'QUANTIDADE_DE_FOLHAS_POR_OF'
               + ';' + sTrackPaginas                                                           // 'QUANTIDADE_DE_PAGINAS_POR_OF'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'CARTAO_POSTAGEM'
               + ';' + sTrackDataLoteQtdPostagem                                               // 'DATA_LOTE_QTD_POSTAGEM'
               + ';' + IntToStr(iTotalLocal)                                                   // 'TOTAL_LOCAL'
               + ';' + IntToStr(iTotalEstadual)                                                // 'TOTAL_ESTADUAL'
               + ';' + IntToStr(iTotalNacional)                                                // 'TOTAL_NACIONAL'
               + ';' + IntToStr(iTotalLocal + iTotalEstadual + iTotalNacional)                 // 'TOTAL'
               + ';' + sTrackPortes                                                            // 'PORTE[GR]'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';'                                                                           // 'FILLER'
               + ';' + sTrackPapel                                                             // 'PAPEL'
               ;

          // TRACK POR ARQUIVO
          ListaOfPorArquivo.Add(sLinha);

      __queryMySQL_processamento__.Next;

    end;

    if not objParametrosDeEntrada.TESTE then
    begin

      //===============================================================================================
      //  MOVE OS REGISTROS DO TRACK_LINE PARA O HISTORY
      //===============================================================================================
      sComando := ' INSERT INTO ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
                + ' ('
                + '   SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                + '   WHERE ARQUIVO_AFP        = "' + sArquivoEntrada + '"'
                +  ' )';
      objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

      sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                + '   WHERE ARQUIVO_AFP        = "' + sArquivoEntrada + '"';
      objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
      //===============================================================================================


      //===============================================================================================
      //  AUALIZA FLAG NA TAELA TRACK
      //===============================================================================================
      sComando := ' UPDATE ' + objParametrosDeEntrada.TABELA_TRACK
                + ' SET STATUS_ARQUIVO = "2"'
                + ' WHERE ARQUIVO_AFP        = "' + sArquivoEntrada + '"'
                +  '  AND STATUS_ARQUIVO = 0';
      objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

      //=======================================================================================================================================
      //  RETIRA EXTENÇÃO PARA CRIAR O ARQUIVO .CSV
      //=======================================================================================================================================
      sArquivoEntrada := StringReplace(sArquivoEntrada, '.AFP', '', [rfReplaceAll, rfIgnoreCase]);

      ListaOfPorArquivo.SaveToFile(sPathCsvTrackFaturamento    + sArquivoEntrada + '.CSV');
      
    end;

    ListaOfPorArquivo.SaveToFile(sPathCsvTrackBKP            + sArquivoEntrada + '.CSV');

    objLogar.Logar(#13 + #10 + ListaOfPorArquivo.Text + #13 + #10);

    ListaOfPorArquivo.Clear;
    //=======================================================================================================================================

  end;

end;

procedure TCore.ExcluirBase(NomeTabela: String);
var
  sComando : String;
  sBase    : string;
begin

  sBase := objString.getTermo(1, '.', NomeTabela);

  sComando := 'drop database ' + sBase;
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
end;

procedure TCore.ExcluirTabela(NomeTabela: String);
var
  sComando : String;
  sTabela  : String;
begin

  sTabela := objString.getTermo(2, '.', NomeTabela);

  sComando := 'drop table ' + sTabela;
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
end;

procedure TCore.StoredProcedure_Dropar(Nome: string; logBD:boolean=false; idprograma:integer=0);
var
  sSQL: string;
  sMensagem: string;
begin
  try
    sSQL := 'DROP PROCEDURE if exists ' + Nome;
    objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1);
  except
    on E:Exception do
    begin
      sMensagem := '  StoredProcedure_Dropar(' + Nome + ') - Excecao:' + E.Message + ' . SQL: ' + sSQL;
      objLogar.Logar(sMensagem);
    end;
  end;

end;

function TCore.StoredProcedure_Criar(Nome : string; scriptSQL: TStringList): boolean;
var
  bExecutou    : boolean;
  sMensagem    : string;
begin


  bExecutou := objConexao.Executar_SQL(__queryMySQL_processamento__, scriptSQL.Text, 1).status;

  if not bExecutou then
  begin
    sMensagem := '  StoredProcedure_Criar(' + Nome + ') - Não foi possível carregar a stored procedure para execução.';
    objLogar.Logar(sMensagem);
  end;

  result := bExecutou;
end;

procedure TCore.StoredProcedure_Executar(Nome: string; ComParametro:boolean=false; logBD:boolean=false; idprograma:integer=0);
var

  sSQL        : string;
  sMensagem   : string;
begin

  try
    (*
    if not Assigned(con) then
    begin
      con := TZConnection.Create(Application);
      con.HostName  := objConexao.getHostName;
      con.Database  := sNomeBase;
      con.User      := objConexao.getUser;
      con.Protocol  := objConexao.getProtocolo;
      con.Password  := objConexao.getPassword;
      con.Properties.Add('CLIENT_MULTI_STATEMENTS=1');
      con.Connected := True;
    end;

    if not Assigned(QP) then
      QP := TZQuery.Create(Application);

    QP.Connection := con;
    QP.SQL.Clear;
    *)

    sSQL := 'CALL '+ Nome;
    if not ComParametro then
      sSQL := sSQL + '()';

    objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1);

  except
    on E:Exception do
    begin
      sMensagem := '[ERRO] StoredProcedure_Executar('+Nome+') - Excecao:'+E.Message+' . SQL: '+sSQL;
      objLogar.Logar(sMensagem);
      ShowMessage(sMensagem);
    end;
  end;

//  objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1)

end;

function TCore.EnviarEmail(Assunto: string=''; Corpo: string=''): Boolean;
var
  sHost    : string;
  suser    : string;
  sFrom    : string;
  sTo      : string;
  sAssunto : string;
  sCorpo   : string;
  sAnexo   : string;
  sAplicacao: string;

begin

  sAplicacao := ExtractFileName(Application.ExeName);
  sAplicacao := StringReplace(sAplicacao, '.exe', '', [rfReplaceAll, rfIgnoreCase]);

  sHost    := objParametrosDeEntrada.eHost;
  suser    := objParametrosDeEntrada.eUser;
  sFrom    := objParametrosDeEntrada.eFrom;
  sTo      := objParametrosDeEntrada.eTo;
  sAssunto := 'Processamento - ' + sAplicacao + ' - ' + objFuncoesWin.GetVersaoDaAplicacao() + ' [PROCESSAMENTO: ' + objParametrosDeEntrada.PEDIDO_LOTE + ']';
  sAssunto := sAssunto + ' ' + Assunto;
  sCorpo   := Corpo;

  sAnexo := objLogar.getArquivoDeLog();

  //sAnexo := StringReplace(anexo, '"', '', [rfReplaceAll, rfIgnoreCase]);
  //sAnexo := StringReplace(anexo, '''', '', [rfReplaceAll, rfIgnoreCase]);

  try

    objEmail := TSMTPDelphi.create(sHost, suser);

    if objEmail.ConectarAoServidorSMTP() then
    begin
      if objEmail.AnexarArquivo(sAnexo) then
      begin

          if not (objEmail.EnviarEmail(sFrom, sTo, sAssunto, sCorpo)) then
            ShowMessage('ERRO AO ENVIAR O E-MAIL')
          else
          if not objEmail.DesconectarDoServidorSMTP() then
            ShowMessage('ERRO AO DESCONECTAR DO SERVIDOR');
      end
      else
        ShowMessage('ERRO AO ANEXAR O ARQUIVO');
    end
    else
      ShowMessage('ERRO AO CONECTAR AO SERVIDOR');

  except
    ShowMessage('NÃO FOI POSSIVEL ENVIAR O E-MAIL.');
  end;
end;



function Tcore.PesquisarLote(LOTE_PEDIDO : STRING; status : Integer): Boolean;
var
  sComando : string;
  iPedido  : Integer;
  sStauts  : string;
begin

  case status of
    0: sStauts := 'S';
    1: sStauts := 'N';
  end;

  objParametrosDeEntrada.PEDIDO_LOTE_TMP := LOTE_PEDIDO;

  sComando := ' SELECT RELATORIO_QTD FROM  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' WHERE LOTE_PEDIDO = ' + LOTE_PEDIDO + ' AND VALIDO = "' + sStauts + '"';
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.stlRelatorioQTDE.Text := __queryMySQL_processamento__.FieldByName('RELATORIO_QTD').AsString;

  if __queryMySQL_processamento__.RecordCount > 0 then
    Result := True
  else
    Result := False;

end;

PROCEDURE TCORE.COMPACTAR();
Var
  sArquivo         : String;
  sPathEntrada     : String;
  sPathSaida       : String;

  iContArquivos    : Integer;
  iTotalDeArquivos : Integer;
BEGIN

  sPathEntrada := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);
  sPathSaida   := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA);
  ForceDirectories(sPathSaida);

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    sArquivo := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivos];
    COMPACTAR_ARQUIVO(sPathEntrada + sArquivo, sPathSaida, True);

  end;

end;

PROCEDURE TCORE.EXTRAIR();
Var
  sArquivo         : String;
  sPathEntrada     : String;
  sPathSaida       : String;

  iContArquivos    : Integer;
  iTotalDeArquivos : Integer;
BEGIN

  sPathEntrada := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);
  sPathSaida   := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA);
  ForceDirectories(sPathSaida);

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    sArquivo := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivos];
    EXTRAIR_ARQUIVO(sPathEntrada + sArquivo, sPathSaida);

  end;

end;


PROCEDURE TCORE.COMPACTAR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String; MOVER_ARQUIVO: Boolean = FALSE);
begin

  Compactar_Arquivo_7z(ARQUIVO_ORIGEM, PATH_DESTINO, MOVER_ARQUIVO);

end;

PROCEDURE TCORE.EXTRAIR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String);
begin

  Extrair_Arquivo_7z(ARQUIVO_ORIGEM, PATH_DESTINO);

end;

function TCORE.Compactar_Arquivo_7z(Arquivo, destino : String; mover_arquivo: Boolean=false): integer;
Var
  sComando                  : String;
  sArquivoDestino           : String;
  sParametros               : String;
  __AplicativoCompactacao__ : String;

  iRetorno                  : Integer;
Begin

    sArquivoDestino := ExtractFileName(Arquivo) + '.7Z';

    destino := objString.AjustaPath(destino);

    sParametros := ' a ';

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 32 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_32bits;

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 64 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_64bits;

    sComando := __AplicativoCompactacao__ + sParametros + ' "' + destino + sArquivoDestino + '" "' + Arquivo + '"';

    if mover_arquivo then
      sComando := sComando + ' -sdel';

    iRetorno := objFuncoesWin.WinExecAndWait32(sComando);

    Result   := iRetorno;

End;

function TCORE.Extrair_Arquivo_7z(Arquivo, destino : String): integer;
Var
  sComando                  : String;
  sParametros               : String;
  __AplicativoCompactacao__ : String;

  iRetorno                  : Integer;
Begin

    destino := objString.AjustaPath(destino);

    sParametros := ' e ';

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 32 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_32bits;

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 64 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_64bits;

    sComando := __AplicativoCompactacao__ + sParametros + ' ' + Arquivo +  ' -y -o"' + destino + '"';

    iRetorno := objFuncoesWin.WinExecAndWait32(sComando);

    Result   := iRetorno;

End;

procedure TCore.CriaMovimento();
var

  txtSaida                          : TextFile;

  sPathEntrada                      : string;

  sPathMovimentoPedido              : string;
  sPathMovimentoIDX                 : string;
  sPathMovimentoAFP                 : string;




  sPathMovimentoArquivos            : string;
  sPathMovimentoBackupZip           : string;
  sPathMovimentoCIF                 : string;
  sPathMovimentoRelatorio           : string;
  sPathComplemento                  : string;
  sPathMovimentoTRACK               : string;
  sPathMovimentoTMP                 : string;
  sArquivoDOC                       : string;
  sArquivoZIP                       : string;
  sArquivoPDF                       : string;
  sArquivoTXT                       : string;
  sArquivoJRN                       : string;
  sArquivoAFP                       : string;
  sArquivoAFP_DOC                   : string;
  sArquivoIDX                       : string;
  sArquivoREL                       : string;
  sComando                          : string;
  sLinha                            : string;

  sDirecao                          : string;
  sCategoria                        : string;
  sPorte                            : string;
  sCep                              : string;

  iContArquivos                     : Integer;
  iContArquivoZip                   : Integer;
  iTotalFolhas                      : Integer;
  iTotalPaginas                     : Integer;
  iTotalObjestos                    : Integer;


  stlFiltroArquivo                  : TStringList;
  stlRelatorio                      : TStringList;
  stlTrack                          : TStringList;

begin

  objParametrosDeEntrada.TIMESTAMP := now();

  stlFiltroArquivo                 := TStringList.create();
  stlRelatorio                     := TStringList.create();
  stlTrack                         := TStringList.create();

  //=======================================================================================================================================================================================
  //  LIMPANDO A TABELA DE PROCESSAMENTO
  //=======================================================================================================================================================================================
  sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO;
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO2;
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
  //=======================================================================================================================================================================================

  if objParametrosDeEntrada.TESTE then
    sPathComplemento := '_TESTE';


  //=======================================================================================================================================================================================
  //  DEFINE ESTRUTURA MOVIMENTO
  //=======================================================================================================================================================================================
  sPathEntrada                     := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);



  //sPathMovimentoArquivos           := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'ARQUIVOS'   + PathDelim;
  //sPathmovimentoBackupZip          := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'BACKUP_ZIP' + PathDelim;
  //sPathmovimentoCIF                := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'CIF'        + PathDelim;
  //sPathMovimentoTRACK              := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'TRACK'      + PathDelim;
  //sPathMovimentoTMP                := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'TMP'      + PathDelim;
  //=======================================================================================================================================================================================

  //===================================================================================================================================================================
  // CRIA PASTAS
  //===================================================================================================================================================================
  //ForceDirectories(sPathmovimentoCIF);
  //ForceDirectories(sPathMovimentoTRACK);
  //ForceDirectories(sPathMovimentoTMP);
  //===================================================================================================================================================================

  //===================================================================================================================================================================
  // ARQUIVOS SELECIONADOS
  //===================================================================================================================================================================
  for iContArquivoZip := 0 to objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Count - 1 do
  begin

    sArquivoAFP := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivoZip];
    sArquivoIDX := StringReplace(sArquivoAFP, '.AFP', '.IDX', [rfReplaceAll, rfIgnoreCase]);
    sArquivoJRN := StringReplace(sArquivoAFP, '.AFP', '.JRN', [rfReplaceAll, rfIgnoreCase]);

    //=======================================================================================================================================
    //  PEGA O NOME DO ARQUIVO ZIP NA TABELA DE TRACK LINE
    //=======================================================================================================================================
    sComando := 'SELECT ARQUIVO_ZIP FROM ' + objParametrosDeEntrada.TABELA_TRACK
              + ' WHERE ARQUIVO_AFP = "' + sArquivoAFP + '" '
              + ' GROUP BY ARQUIVO_ZIP';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

    sArquivoZIP := __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString;
    //=======================================================================================================================================


    //=======================================================================================================================================
    //  CARREGA O IDX NA TABELA PROCESSAMENTO2
    //=======================================================================================================================================
    sComando := ' LOAD DATA LOCAL INFILE "' + StringReplace(sPathEntrada, '\', '\\', [rfReplaceAll, rfIgnoreCase]) + sArquivoIDX + '" '
             + '  INTO TABLE ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO2
             + '    CHARACTER SET latin1 '
             + '  LINES '
             + '    TERMINATED BY "\r\n" '
             + '   SET SEQUENCIA      = MID(LINHA, 1, 8) '
             + '      ,ARQUIVO_ZIP    = "' + sArquivoZIP + '"'
             + '      ,ARQUIVO_AFP    = "' + sArquivoAFP + '"'
             + '      ,MOVIMENTO      = "' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + '"'
             ;
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
    //=======================================================================================================================================

    //=======================================================================================================================================
    //  CARREGA O JRN NA TABELA PROCESSAMENTO
    //=======================================================================================================================================
    sComando := ' LOAD DATA LOCAL INFILE "' + StringReplace(sPathEntrada, '\', '\\', [rfReplaceAll, rfIgnoreCase]) + sArquivoJRN + '" '
             + '  INTO TABLE ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
             + '    CHARACTER SET latin1 '
             + '  FIELDS '
             + '    TERMINATED BY "|" '
             + '  LINES '
             + '    TERMINATED BY "\r\n" '
             + '   SET LOTE          = MID(CIF, 11, 5) '
             + '      ,DATA_POSTAGEM = MID(CIF, 29, 6) '
             + '      ,ARQUIVO_AFP   = "' + sArquivoAFP + '"'
             + '      ,ARQUIVO_ZIP   = "' + sArquivoZIP + '"'
             + '      ,MOVIMENTO     = "' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + '"'
             + '      ,OF_FORMULARIO    = "' + objParametrosDeEntrada.OF_FORMULARIO + '"'
             + '      ,PESO             = "' + objParametrosDeEntrada.PESO_PAPEL    + '"'
             + '      ,ACABAMENTO       = "' + objParametrosDeEntrada.ACABAMENTO    + '"'
             + '      ,PAPEL            = "' + objParametrosDeEntrada.PAPEL         + '"'
             + '      ,INDICE_CEP_PLANO = ('
             + '                             SELECT SEQ FROM ' + objParametrosDeEntrada.TABELA_PLANO_DE_TRIAGEM
             + '                             WHERE CEPINI <= MID(CEP, 2, 8) AND CEPFIN >= MID(CEP, 2, 8) '
             + ')'
             + '      ,ARQUIVO_COUNT    = "' + IntToStr(iContArquivos+1) + '"'
             ;
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
    //=======================================================================================================================================

  end;

  //=======================================================================================================================================
  //  FAZ A SEGMENTAÇÃO DE SAÍDA
  //=======================================================================================================================================
  sComando := 'SELECT OF_FORMULARIO, ACABAMENTO, PAPEL, ARQUIVO_COUNT, ARQUIVO_AFP, DATA_POSTAGEM FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
            + ' WHERE INDICE_CEP_PLANO IS NOT NULL '
            + ' GROUP BY OF_FORMULARIO, ACABAMENTO, PAPEL, ARQUIVO_COUNT, DATA_POSTAGEM';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  sPathMovimentoPedido := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim;

  while not __queryMySQL_processamento__.Eof do
  begin

    //================================================================================================================================
    // CRIANDO IDX ARQUIVOS DE SAÍDA
    //================================================================================================================================
    sComando := ' SELECT IDX.LINHA, JRN.* FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO2 + ' AS IDX '
              + '                    LEFT JOIN ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO  + ' AS JRN '
              + '  ON IDX.SEQUENCIA = JRN.SEQUENCIA '
              + '  WHERE JRN.INDICE_CEP_PLANO IS NOT NULL '
              + '    AND JRN.OF_FORMULARIO = "' + __queryMySQL_processamento__.FieldByName('OF_FORMULARIO').AsString + '" '
              + '    AND JRN.ACABAMENTO    = "' + __queryMySQL_processamento__.FieldByName('ACABAMENTO').AsString + '" '
              + '    AND JRN.PAPEL         = "' + __queryMySQL_processamento__.FieldByName('PAPEL').AsString + '" '
              + '    AND JRN.ARQUIVO_COUNT =  ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_COUNT').AsString
              + '    AND JRN.DATA_POSTAGEM =  ' + __queryMySQL_processamento__.FieldByName('DATA_POSTAGEM').AsString
              + '  ORDER BY JRN.INDICE_CEP_PLANO ';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

    if __queryMySQL_processamento2__.RecordCount > 0 then
    begin

        //==============================================================================
        // ARQUIVO AFP ORIGEM
        //==============================================================================
        sArquivoAFP := __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString;
        //==============================================================================

        //==============================================================================================================================================================================================================================================================================================
        // CRIA A PASTA DA SAIDA POR OF POR ARQUIVO DE ENTRADA
        //==============================================================================================================================================================================================================================================================================================
        sPathMovimentoIDX                := sPathMovimentoPedido + 'POSTAGEN_' + __queryMySQL_processamento__.FieldByName('DATA_POSTAGEM').AsString + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, __queryMySQL_processamento__.FieldByName('ARQUIVO_COUNT').AsInteger) + PathDelim + __queryMySQL_processamento__.FieldByName('OF_FORMULARIO').AsString + '_' + __queryMySQL_processamento__.FieldByName('ACABAMENTO').AsString + '_' + __queryMySQL_processamento__.FieldByName('PAPEL').AsString + PathDelim;
        sPathMovimentoAFP                := sPathMovimentoIDX + 'AFP' + PathDelim;

        ForceDirectories(sPathMovimentoAFP);
        //==============================================================================================================================================================================================================================================================================================

        sArquivoDOC     := 'DOC_' + __queryMySQL_processamento__.FieldByName('OF_FORMULARIO').AsString + '_' + __queryMySQL_processamento__.FieldByName('ACABAMENTO').AsString + '_' + __queryMySQL_processamento__.FieldByName('PAPEL').AsString + '_' + __queryMySQL_processamento__.FieldByName('ARQUIVO_COUNT').AsString + '.IDX';
        sArquivoAFP_DOC := StringReplace(sArquivoDOC, '.IDX', '.AFP', [rfReplaceAll, rfIgnoreCase]);

        AssignFile(txtSaida, sPathMovimentoIDX + sArquivoDOC);
        Rewrite(txtSaida);


        while NOT __queryMySQL_processamento2__.Eof DO
        begin

          sLinha := objString.AjustaStr(FormatFloat('00000000', __queryMySQL_processamento2__.RecNo), 8)
                   +  __queryMySQL_processamento2__.FieldByName('LINHA').AsString
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('AUDIT').AsString, 10)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('CIF').AsString, 35)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('PAGINAS').AsString, 4)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FOLHAS').AsString, 4)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('PAGINA_INICIAL').AsString, 7)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('PAGINA_FINAL').AsString, 7)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('NOME').AsString, 50)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('LOGRADOURO').AsString, 110)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('CEP').AsString, 10)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_01').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_02').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_03').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_04').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_05').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('CODIGO_BARRAS').AsString, 50)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_06').AsString, 2)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('FILLER_07').AsString, 6)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('NOME_2').AsString, 50)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('DEVOLUCAO').AsString, 10)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('DATA_POSTAGEM').AsString, 10)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('ARQUIVO_AFP').AsString, 50)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('ARQUIVO_ZIP').AsString, 50)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('MOVIMENTO').AsString, 10)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('OF_FORMULARIO').AsString, 11)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('PESO').AsString, 7)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('ACABAMENTO').AsString, 20)
                   + objString.AjustaStr(__queryMySQL_processamento2__.FieldByName('PAPEL').AsString, 10)
                   + FormatFloat('0000000', __queryMySQL_processamento2__.FieldByName('INDICE_CEP_PLANO').AsInteger);

          writeln(txtSaida, slinha);

          //====================================================================
          //                            DEFINE O PORTE
          //====================================================================
          case __queryMySQL_processamento2__.FieldByName('PESO').AsInteger of
            0001..2000: sPorte := '1';
            2001..5000: sPorte := '2';
          else
             sPorte := '3';
          end;
          //====================================================================

          //==============================================================
          //  FAC NORMAL
          //==============================================================
          sCep := Copy(__queryMySQL_processamento2__.FieldByName('CEP').AsString, 2, 8);

          if sCEP < '10000000' then  //Apurando o Destino/Categoria
          begin

            sDirecao   := '1';
            sCategoria := '82015'; // Grande S.Paulo //

          end
          else
          begin

            if sCEP < '20000000' then
            begin

              sDirecao   := '2';
              sCategoria := '82023'; // Interior de S.Paulo //

            end
            else
            begin

              sDirecao   := '3';
              sCategoria := '82031'; // Outros Estados //

            end

          end;
          //======================================================================================================================================================

          //=================================================================================================================================================================
          //  INSERE NA TABELA TRACK E CRIA CSV TRACK PRÉVIAS
          //=================================================================================================================================================================
          if not objParametrosDeEntrada.TESTE then
          begin
            sComando := 'INSERT INTO  ' + objParametrosDeEntrada.TABELA_TRACK_LINE
                      + ' (ARQUIVO_ZIP'
                       + ',ARQUIVO_AFP'
                       + ',SEQUENCIA_REGISTRO'
                       + ',TIMESTAMP'
                       + ',LOTE_PROCESSAMENTO'
                       + ',MOVIMENTO'
                       + ',ACABAMENTO'
                       + ',PAGINAS'
                       + ',FOLHAS'
                       + ',OF_FORMULARIO'
                       + ',DATA_POSTAGEM'
                       + ',LOTE'
                       + ',CIF'
                       + ',PESO'
                       + ',DIRECAO'
                       + ',CATEGORIA'
                       + ',PORTE'
                       + ',STATUS_REGISTRO'
                       + ',PAPEL'
                       + ') '
                       + ' VALUES("'
                       +         __queryMySQL_processamento2__.FieldByName('ARQUIVO_ZIP').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('ARQUIVO_AFP').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('SEQUENCIA').AsString
                       + '","' + FormatDateTime('YYYY-MM-DD hh:mm:ss', objParametrosDeEntrada.TIMESTAMP)
                       + '","' + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE))
                       + '","' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO)
                       + '","' + __queryMySQL_processamento2__.FieldByName('ACABAMENTO').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('PAGINAS').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('FOLHAS').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('OF_FORMULARIO').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('DATA_POSTAGEM').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('LOTE').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('CIF').AsString
                       + '","' + __queryMySQL_processamento2__.FieldByName('PESO').AsString
                       + '","' + sDirecao
                       + '","' + sCategoria
                       + '","' + sPorte
                       + '","' + '0'
                       + '","' + __queryMySQL_processamento2__.FieldByName('PAPEL').AsString
                       + '")'
                       ;
            objConexao.Executar_SQL(__queryMySQL_Insert_, sComando, 1);

          end;

          __queryMySQL_processamento2__.Next
        end;

        CloseFile(txtSaida);

        //===============================================================================================================================================
        // COPIANDO AFP DE ENTRADA TEMPORÁRIO PARA CRIAÇÃO DO SPOOL
        //===============================================================================================================================================
        objFuncoesWin.CopiarArquivo(sPathEntrada + sArquivoAFP, sPathMovimentoIDX + sArquivoAFP_DOC);
        //===============================================================================================================================================

        //===================================================================================================
        // CRIANDO SPOOL AFP PARA O ARQUIVO IDX GERADO
        //===============================================================================================================================================
        Atualiza_arquivo_conf_C(objParametrosDeEntrada.APP_C_GERA_SPOOL_CFG, sPathMovimentoIDX, sPathMovimentoAFP, '', '', '');
        execulta_app_c(objParametrosDeEntrada.APP_C_GERA_SPOOL_EXE, objParametrosDeEntrada.APP_C_GERA_SPOOL_CFG);
        //===============================================================================================================================================

        //===============================================================================================================================================
        // EXCLUI O ARQUIVO AFP ENTRADA APÓS GERAÇÃO DO NOVO SPOOL
        //===============================================================================================================================================
        DeleteFile(sPathMovimentoIDX + sArquivoAFP_DOC);
        //===============================================================================================================================================

    end;
    //================================================================================================================================

    __queryMySQL_processamento__.Next;
  end;

  //==================================================================================================================================================================================================
  // CRIANDO RELATÓRIO DE QUANTIDADES
  //==================================================================================================================================================================================================
  stlRelatorio.Clear;
  sComando := 'SELECT '
            + '  concat(mid(MOVIMENTO, 1, 4), "-", mid(MOVIMENTO, 5, 2), "-", mid(MOVIMENTO, 7, 2)) as MOVIMENTO'
            + ', concat(mid(DATA_POSTAGEM, 1, 2), "/", mid(DATA_POSTAGEM, 3, 2), "/", mid(DATA_POSTAGEM, 5, 2)) as DATA_POSTAGEM'
            + ', LOTE'
            + ', OF_FORMULARIO'
            + ', ACABAMENTO'
            + ', PAPEL'
            + ', COUNT(OF_FORMULARIO) AS QUANTIDADE '
            + ', SUM(PAGINAS)         AS PAGINAS '
            + ', SUM(FOLHAS)          AS FOLHAS '
            + ' FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
            + ' WHERE ARQUIVO_COUNT IS NOT NULL and DATA_POSTAGEM <> "" '
            + ' GROUP BY OF_FORMULARIO, ACABAMENTO, PAPEL, LOTE';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  sLinha := stringOfChar('-', 106)
  + #13 + #10 + 'RELATÓRIO DE QUANTIDADES - PROCESSAMENTO'
  + #13 + #10 + stringOfChar('-', 106)
  + #13 + #10 + 'MOVIMENTO  DATA POST. LOTE POST. OF FORMULARIO      ACABAMENTO      PAPEL QUANTIDADE    PAGINAS     FOLHAS'
  + #13 + #10 + '---------- ---------- ---------- ------------- --------------- ---------- ---------- ---------- ----------';

  stlRelatorio.Add(sLinha);

  iTotalObjestos  := 0;
  iTotalFolhas    := 0;
  iTotalPaginas   := 0;

  while not __queryMySQL_processamento__.Eof do
  begin

    sLinha := objString.AjustaStr(__queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString, 10, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('DATA_POSTAGEM').AsString, 10, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('LOTE').AsString, 10, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('OF_FORMULARIO').AsString, 13, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('ACABAMENTO').AsString, 15, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('PAPEL').AsString, 10, 1)
      + ' ' + FormatFloat('0000000000', __queryMySQL_processamento__.FieldByName('QUANTIDADE').AsInteger)
      + ' ' + FormatFloat('0000000000',    __queryMySQL_processamento__.FieldByName('PAGINAS').AsInteger)
      + ' ' + FormatFloat('0000000000',    __queryMySQL_processamento__.FieldByName('FOLHAS').AsInteger)
      ;
    stlRelatorio.Add(sLinha);

    //=================================================================================================================================================================

    iTotalObjestos  := iTotalObjestos  + __queryMySQL_processamento__.FieldByName('QUANTIDADE').AsInteger;
    iTotalFolhas    := iTotalFolhas    + __queryMySQL_processamento__.FieldByName('FOLHAS').AsInteger;
    iTotalPaginas   := iTotalPaginas   + __queryMySQL_processamento__.FieldByName('PAGINAS').AsInteger;

    __queryMySQL_processamento__.Next;
  end;

  sLinha := stringOfChar('-', 73) + ' ---------- ---------- ----------'
  + #13 + #10 + 'TOTAIS' + stringOfChar(' ', 68) + FormatFloat('0000000000', iTotalObjestos) + ' ' + FormatFloat('0000000000', iTotalPaginas) + ' ' + FormatFloat('0000000000', iTotalFolhas);
  stlRelatorio.Add(sLinha);

  sArquivoREL := sPathMovimentoPedido + 'RELATORIO_DE_QUANTIDADES_' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) +'.REL';

  stlRelatorio.SaveToFile(sArquivoREL);
  objLogar.Logar(#13 + #10 + stlRelatorio.Text + #13 + #10);

  objFuncoesWin.ExecutarArquivoComProgramaDefault(sArquivoREL);
  //==================================================================================================================================================================================================


  //==================================================================================================================================================
  //  ATUALIZANDO STATUS DO ARQUIVO NA TABELA TRACK
  //==================================================================================================================================================
  sComando := 'SELECT ARQUIVO_ZIP FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
            + ' GROUP BY ARQUIVO_AFP';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  while NOT __queryMySQL_processamento__.Eof DO
  BEGIN

    sComando := 'UPDATE ' + objParametrosDeEntrada.TABELA_TRACK
              + ' SET STATUS_ARQUIVO = 1'
              + ' WHERE ARQUIVO_ZIP = "' + __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString + '"';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 1);

    __queryMySQL_processamento__.Next;
  end;

  //==================================================================================================================================================


  (*
  __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString;
  //=======================================================================================================================================






  //===================================================================================================================================================================
  // EXTRAI E MOVE OS ARQUIVOS
  //===================================================================================================================================================================
  for iContArquivoZip := 0 to objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Count - 1 do
  begin

    objFuncoesWin.DeletarArquivosPorFiltro(sPathMovimentoTMP , '*.*');

    sArquivoZIP  := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivoZip];

    EXTRAIR_ARQUIVO(sPathEntrada + sArquivoZIP, sPathMovimentoTMP);

    objFuncoesWin.CopiarArquivo(sPathEntrada + sArquivoZIP, sPathmovimentoBackupZip + sArquivoZIP);
    objFuncoesWin.CopiarArquivo(sPathEntrada + sArquivoZIP, sPathMovimentoTMP       + sArquivoZIP);
    DeleteFile(sPathMovimentoTMP + sArquivoZIP);

    //===================================================================================================================================================================
    // MOVENDO OS ARQUIVOS
    //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS PDF (LISTA DE POSTAGEM) E MOVE PAA A PASTA DE POSTAGEM
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.PDF');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoPDF := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoPDF, sPathmovimentoCIF + sArquivoPDF) then
         DeleteFile(sPathMovimentoTMP + sArquivoPDF);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (CIF) E MOVE PAA A PASTA DE POSTAGEM
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.TXT');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoTXT := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoTXT, sPathmovimentoCIF + sArquivoTXT) then
         DeleteFile(sPathMovimentoTMP + sArquivoTXT);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (AFP) E MOVE PAA A PASTA DE ARQUIVOS
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.AFP');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoAFP := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoAFP, sPathMovimentoArquivos + sArquivoAFP) then
         DeleteFile(sPathMovimentoTMP + sArquivoAFP);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (JRN) E MOVE PAA A PASTA DE ARQUIVOS
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.JRN');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoJRN := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoJRN, sPathMovimentoArquivos + sArquivoJRN) then
         DeleteFile(sPathMovimentoTMP + sArquivoJRN);
      end;
      //===================================================================================================================================================================

    //===================================================================================================================================================================

    //===================================================================================================================================================================
    // CARREGA ARQUIVO JRN PARA BANCO PARA GERAR RELATÓRIOS
    //===================================================================================================================================================================
    stlFiltroArquivo.Clear;
    objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoArquivos, stlFiltroArquivo, '*.JRN');
    for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
    begin

      sArquivoJRN := stlFiltroArquivo.Strings[iContArquivos];
      sArquivoAFP := StringReplace(sArquivoJRN, '.JRN', '.AFP', [rfReplaceAll, rfIgnoreCase]);

      sComando := ' LOAD DATA LOCAL INFILE "' + StringReplace(sPathMovimentoArquivos, '\', '\\', [rfReplaceAll, rfIgnoreCase]) + sArquivoJRN + '" '
               + '  INTO TABLE ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
               + '    CHARACTER SET latin1 '
               + '  FIELDS '
               + '    TERMINATED BY "|" '
               + '  LINES '
               + '    TERMINATED BY "\r\n" '
               + '   SET LOTE          = MID(CIF, 11, 5) '
               + '      ,DATA_POSTAGEM = MID(CIF, 29, 6) '
               + '      ,ARQUIVO_AFP   = "' + sArquivoAFP + '"'
               + '      ,ARQUIVO_ZIP   = "' + sArquivoZIP + '"'
               + '      ,MOVIMENTO     = "' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + '"'
               ;
      objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
    end;
    //===================================================================================================================================================================



  END;
  //===================================================================================================================================================================


  //===================================================================================================
  // CABEÇALHO DO CSV TRACK
  //==================================================================================================================================================================
  stlTrack.Clear;
  sLinha      :=  'OF_FORMULARIO'
               + ';FILLER'
               + ';FILLER'
               + ';MOVIMENTO'
               + ';FILLER'
               + ';ARQUIVO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';TIMESTAMP'
               + ';ACABAMENTO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';LOTE_PROCESAMENTO'
               + ';QUANTIDADE_DE_OBJETOS_POR_OF'
               + ';QUANTIDADE_DE_FOLHAS_POR_OF'
               + ';QUANTIDADE_DE_PAGINAS_POR_OF'
               + ';FILLER'
               + ';FILLER'
               + ';CARTAO_POSTAGEM'
               + ';DATA_LOTE_QTD_POSTAGEM'
               + ';TOTAL_LOCAL'
               + ';TOTAL_ESTADUAL'
               + ';TOTAL_NACIONAL'
               + ';TOTAL'
               + ';PORTE[GR]'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';PAPEL'
               ;
    stlTrack.Add(sLinha);
  //==================================================================================================================================================================


  objFuncoesWin.DelTree(sPathMovimentoTMP);
  //==================================================================================================================================================================
  *)

  objLogar.Logar('');


end;

procedure TCore.Atualiza_arquivo_conf_C(ArquivoConf, sINP, sOUT, sTMP, sLOG, sRGP: String);
var
  txtEntrada       : TextFile;
  sLinha           : string;
  sParametro       : string;
  stlArquivoConfC  : TStringList;
  sPathSaidaAFP    : string;
begin


  stlArquivoConfC := TStringList.Create();

  AssignFile(txtEntrada, ArquivoConf);
  Reset(txtEntrada);

  while not Eof(txtEntrada) do
  begin

    Readln(txtEntrada, sLinha);

    sParametro := AnsiUpperCase(Trim(objString.getTermo(1, '=', sLinha)));

    if sParametro = 'INP' then
      stlArquivoConfC.Add(sParametro + '=' + sINP);

    if sParametro = 'OUT' then
      stlArquivoConfC.Add(sParametro + '=' + sOUT);

    if sParametro = 'TMP' then
      stlArquivoConfC.Add(sParametro + '=' + sTMP);

    if sParametro = 'LOG' then
      stlArquivoConfC.Add(sParametro + '=' + sLOG);

    if sParametro = 'RGP' then
      stlArquivoConfC.Add(sParametro + '=' + sRGP);

  end;

  CloseFile(txtEntrada);

  stlArquivoConfC.SaveToFile(ArquivoConf);

end;

procedure TCore.execulta_app_c(app, arquivo_conf: string);
begin
  objFuncoesWin.ExecutarPrograma(app + ' "' + arquivo_conf + '"');
end;

function TCore.ArquivoExieteTabelaTrackLine(Arquivo: string): Boolean;
var
  sComando: string;
begin

  sComando := 'SELECT ARQUIVO_AFP FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
            + ' WHERE ARQUIVO_AFP = "' + Arquivo + '" ';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  if __queryMySQL_processamento__.RecordCount > 0 then
   Result := True
  else
    Result := False;

end;

procedure TCore.getListaDeArquivosPendentes();
var
  sComando                   : string;
  sArquivoAFP                : string;
  sLinha                     : string;
  bPendente                  : Boolean;
begin

  sComando := ' SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK
            + ' WHERE STATUS_ARQUIVO = 0 '
            + '    OR STATUS_ARQUIVO = 1 '
            + ' GROUP BY ARQUIVO_AFP '
            + ' ORDER BY MOVIMENTO DESC, ARQUIVO_AFP ';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.STL_LISTA_ARQUIVOS_PENDENTES.Clear;

  WHILE NOT __queryMySQL_processamento__.Eof do
  BEGIN

    sArquivoAFP := __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString;

    bPendente := True;
    //====================================================================================
    //  VERIFICA SE O ARUQIVO ESTÁ NA TABELA TRACK_LINE
    //====================================================================================
    sComando := ' SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
              + ' WHERE ARQUIVO_AFP = "' + sArquivoAFP + '"'
              + ' group by ARQUIVO_AFP';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

    if __queryMySQL_processamento2__.RecordCount > 0 then
      bPendente := False;
    //====================================================================================

    //====================================================================================
    //  VERIFICA SE O ARUQIVO ESTÁ NA TABELA TRACK_LINE
    //====================================================================================
    sComando := ' SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
              + ' WHERE ARQUIVO_AFP = "' + sArquivoAFP + '"'
              + ' group by ARQUIVO_AFP';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

    if __queryMySQL_processamento2__.RecordCount > 0 then
      bPendente := False;
    //====================================================================================

    IF bPendente THEN
    BEGIN
      sLinha := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString
      + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString
      + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString;

      objParametrosDeEntrada.STL_LISTA_ARQUIVOS_PENDENTES.Add(sLinha);
    end;

    __queryMySQL_processamento__.Next;
  end;

end;

procedure TCore.getListaDeArquivosTrack();
var
  sComando                   : string;
  sArquivoAFP                : string;
  sLinha                     : string;
  bPendente                  : Boolean;
begin

  sComando := ' SELECT MOVIMENTO, ARQUIVO_AFP, ARQUIVO_ZIP, COUNT(ARQUIVO_AFP) AS QTD FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE
            + ' GROUP BY ARQUIVO_AFP '
            + ' ORDER BY MOVIMENTO DESC, ARQUIVO_AFP ';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.STL_LISTA_ARQUIVOS_TRACK.Clear;

  WHILE NOT __queryMySQL_processamento__.Eof do
  BEGIN

    sArquivoAFP := __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString;

    bPendente := False;
    //====================================================================================
    //  VERIFICA SE O ARUQIVO ESTÁ NA TABELA TRACK_LINE
    //====================================================================================
    sComando := ' SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
              + ' WHERE ARQUIVO_AFP = "' + sArquivoAFP + '"'
              + ' group by ARQUIVO_AFP';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

    if __queryMySQL_processamento2__.RecordCount > 0 then
      bPendente := True;
    //====================================================================================

    //====================================================================================
    //  VERIFICA SE O ARUQIVO ESTÁ NA TABELA TRACK_LINE
    //====================================================================================
    sComando := ' SELECT ARQUIVO_AFP, OBJETOS FROM ' + objParametrosDeEntrada.TABELA_TRACK
              + ' WHERE ARQUIVO_AFP = "' + sArquivoAFP + '"'
              + ' group by ARQUIVO_AFP';
    objConexao.Executar_SQL(__queryMySQL_processamento2__, sComando, 2);

    if __queryMySQL_processamento__.FieldByName('QTD').AsInteger <> __queryMySQL_processamento2__.FieldByName('OBJETOS').AsInteger then
      bPendente := True;
    //====================================================================================

    IF NOT bPendente THEN
    BEGIN
      sLinha := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString
      + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString
      + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString;

      objParametrosDeEntrada.STL_LISTA_ARQUIVOS_TRACK.Add(sLinha);
    end;

    __queryMySQL_processamento__.Next;
  end;

end;

procedure TCore.getListaDeArquivosJaProcessados();
var
  sComando                   : string;
  sLinha                     : string;
begin

  sComando := ' SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
            + ' GROUP BY ARQUIVO_AFP '
            + ' ORDER BY MOVIMENTO DESC ';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS.Clear;

  WHILE NOT __queryMySQL_processamento__.Eof do
  BEGIN

    sLinha := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString
    + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString;

    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS.Add(sLinha);

    __queryMySQL_processamento__.Next;
  end;

end;

procedure TCore.ReverterArquivos();
var
  iContArquivos                       : Integer;
  sArquivoReverter                    : string;
  sComando                            : string;

begin

  for iContArquivos := 0 to objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER.Count - 1 do
  begin

    sArquivoReverter := objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER.Strings[iContArquivos];

    //===============================================================================================
    //  MOVE OS REGISTROS DO TRACK_LINE PARA O HISTORY
    //===============================================================================================
    sComando := ' INSERT INTO ' + objParametrosDeEntrada.TABELA_TRACK_LINE
              + ' ('
              + '   SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
              + '   WHERE ARQUIVO_AFP        = "' + sArquivoReverter + '"'
              +  ' )';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

    sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY
              + ' WHERE ARQUIVO_AFP = "' + sArquivoReverter + '" ';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

    sComando := 'UPDATE ' + objParametrosDeEntrada.TABELA_TRACK
              + ' SET STATUS_ARQUIVO = 1'
              + ' WHERE ARQUIVO_AFP = "' + sArquivoReverter + '" ';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  end;

end;

end.
