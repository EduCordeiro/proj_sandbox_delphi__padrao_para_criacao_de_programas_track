CREATE DATABASE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track;

DROP TABLE IF EXISTS       proj_sandbox_delphi__padrao_para_criacao_de_programas_track.processamento;
CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.processamento (
   SEQUENCIA                  INTEGER
  ,AUDIT                      varchar(008) default NULL
  ,CIF                        varchar(034) default NULL
  ,PAGINAS                    INTEGER
  ,FOLHAS                     INTEGER
  ,PAGINA_INICIAL             INTEGER
  ,PAGINA_FINAL               INTEGER
  ,NOME                       varchar(040) default NULL
  ,LOGRADOURO                 varchar(100) default NULL
  ,CEP                        varchar(009) default NULL
  ,FILLER_01                  varchar(001) default NULL
  ,FILLER_02                  varchar(001) default NULL
  ,FILLER_03                  varchar(001) default NULL
  ,FILLER_04                  varchar(001) default NULL
  ,FILLER_05                  varchar(001) default NULL
  ,CODIGO_BARRAS              varchar(044) default NULL
  ,FILLER_06                  varchar(001) default NULL
  ,FILLER_07                  varchar(005) default NULL
  ,NOME_2                     varchar(040) default NULL
  ,DEVOLUCAO                  varchar(010) default NULL
  ,LOTE                       varchar(005) default NULL
  ,DATA_POSTAGEM              varchar(006) default NULL
  ,ARQUIVO_AFP                varchar(050) default NULL
  ,ARQUIVO_ZIP                varchar(050) default NULL
  ,MOVIMENTO                  varchar(008) default NULL
  ,OF_FORMULARIO              varchar(010) default NULL
  ,PESO                       INTEGER
  ,ACABAMENTO                 varchar(020) default NULL
  ,PAPEL                      varchar(010) default NULL
  ,INDICE_CEP_PLANO           INTEGER
  ,ARQUIVO_COUNT              INTEGER
  ,PRIMARY KEY(SEQUENCIA)
);

/*DROP TABLE IF EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.controle_arquivos;*/
CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.controle_arquivos (
  LOTE                 int(10)      unsigned NOT NULL,
  DATA_INSERSAO        datetime              NOT NULL,
  ARQUIVO              varchar(100)          NOT NULL,
  PAGINAS              varchar(010)          NOT NULL,
  OBJETOS              varchar(010)          NOT NULL,
  PRIMARY KEY (LOTE, ARQUIVO),
  KEY idx_controle_arquivo (ARQUIVO)
);

CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.LOTES_PEDIDOS (
  LOTE_PEDIDO      int     NOT NULL auto_increment,
  VALIDO           CHAR(1) NOT NULL default 'N',

  DATA_CRIACAO     DATETIME,
  CHAVE            VARCHAR(17),
  ID               VARCHAR(17),
  USUARIO_WIN      VARCHAR(20),
  USUARIO_APP      VARCHAR(20),
  IP               VARCHAR(14),
  LOTE_LOGIN       INT,

  RELATORIO_QTD    MEDIUMBLOB,
  HOSTNAME         varchar(15),
  PRIMARY KEY  (LOTE_PEDIDO)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
DROP TABLE IF EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.tbl_blocagem;

CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.track (
  ARQUIVO_ZIP              VARCHAR(60) NOT NULL,
  ARQUIVO_AFP              VARCHAR(60) NOT NULL,
  LOTE                     INT(11) NOT NULL,
  TIMESTAMP                DATETIME NOT NULL,
  LINHAS                   INT(11) NOT NULL DEFAULT '0',
  OBJETOS                  INT(11) NOT NULL DEFAULT '0',
  FOLHAS                   INT(11) NOT NULL DEFAULT '0',
  PAGINAS                  INT(11) NOT NULL DEFAULT '0',
  CEP_INVALIDO             INT(11) NOT NULL DEFAULT '0',
  CEP_VALIDO               INT(11) NOT NULL DEFAULT '0',
  STATUS_ARQUIVO           INT(11) NOT NULL DEFAULT '0',
  MOVIMENTO                VARCHAR(8) NOT NULL,
  PRIMARY KEY  (ARQUIVO_ZIP, ARQUIVO_AFP)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;


CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.track_line (
  ARQUIVO_ZIP            VARCHAR(60) NOT NULL,
  ARQUIVO_AFP            VARCHAR(60) NOT NULL,
  SEQUENCIA_REGISTRO     INT(11) NOT NULL,
  TIMESTAMP              DATETIME NOT NULL,
  LOTE_PROCESSAMENTO     INT(11) NOT NULL,
  MOVIMENTO              VARCHAR(8) NOT NULL,
  ACABAMENTO             VARCHAR(20) NOT NULL,
  PAGINAS                INT(11) NOT NULL DEFAULT '0',
  FOLHAS                 INT(11) NOT NULL DEFAULT '0',
  ENCARTES               INT(11) NOT NULL DEFAULT '0',
  OF_ENVELOPE            VARCHAR(15) NOT NULL,
  OF_FORMULARIO          VARCHAR(15) NOT NULL,
  DATA_POSTAGEM          VARCHAR(10) NOT NULL,
  LOTE                   VARCHAR(5) NOT NULL,
  CARTAO                 VARCHAR(12) NOT NULL,
  CIF                    VARCHAR(34) NOT NULL,
  PESO                   VARCHAR(10) NOT NULL,
  DIRECAO                INT(11) NOT NULL,
  CATEGORIA              INT(11) NOT NULL,
  PORTE                  INT(11) NOT NULL,
  STATUS_REGISTRO        VARCHAR(20) NOT NULL,
  PAPEL                  VARCHAR(10) NOT NULL,
  PRIMARY KEY  (ARQUIVO_ZIP,ARQUIVO_AFP,SEQUENCIA_REGISTRO)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS proj_sandbox_delphi__padrao_para_criacao_de_programas_track.track_line_history (
  ARQUIVO_ZIP            VARCHAR(60) NOT NULL,
  ARQUIVO_AFP            VARCHAR(60) NOT NULL,
  SEQUENCIA_REGISTRO     INT(11) NOT NULL,
  TIMESTAMP              DATETIME NOT NULL,
  LOTE_PROCESSAMENTO     INT(11) NOT NULL,
  MOVIMENTO              VARCHAR(8) NOT NULL,
  ACABAMENTO             VARCHAR(20) NOT NULL,
  PAGINAS                INT(11) NOT NULL DEFAULT '0',
  FOLHAS                 INT(11) NOT NULL DEFAULT '0',
  ENCARTES               INT(11) NOT NULL DEFAULT '0',
  OF_ENVELOPE            VARCHAR(15) NOT NULL,
  OF_FORMULARIO          VARCHAR(15) NOT NULL,
  DATA_POSTAGEM          VARCHAR(10) NOT NULL,
  LOTE                   VARCHAR(5) NOT NULL,
  CARTAO                 VARCHAR(12) NOT NULL,
  CIF                    VARCHAR(34) NOT NULL,
  PESO                   VARCHAR(10) NOT NULL,
  DIRECAO                INT(11) NOT NULL,
  CATEGORIA              INT(11) NOT NULL,
  PORTE                  INT(11) NOT NULL,
  STATUS_REGISTRO        VARCHAR(20) NOT NULL,
  PAPEL                  VARCHAR(10) NOT NULL,
  PRIMARY KEY  (ARQUIVO_ZIP,ARQUIVO_AFP,SEQUENCIA_REGISTRO)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
