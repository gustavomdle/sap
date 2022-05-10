import * as React from 'react';
import styles from './SapDetalhesProposta.module.scss';
import { ISapDetalhesPropostaProps } from './ISapDetalhesPropostaProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jquery from 'jquery';
import * as $ from "jquery";
import * as jQuery from "jquery";
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import "bootstrap";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { Web } from "sp-pnp-js";
import pnp from "sp-pnp-js";
import { ICamlQuery } from '@pnp/sp/lists';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement, DatePicker } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/files";

import InputMask from 'react-input-mask';
import { deprecationHandler } from 'moment';
import { dateToNumber } from '@pnp/spfx-controls-react';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _idProposta;
var _web;


export default class SapDetalhesProposta extends React.Component<ISapDetalhesPropostaProps, {}> {


  public componentDidMount() {

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idProposta = parseInt(queryParms.getValue("PropostasID"));

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    this.getProposta();
    this.getTarefas();
    this.getAnexos();

  }



  public render(): React.ReactElement<ISapDetalhesPropostaProps> {
    return (


      <div id="container">

        <div id="accordion">

          <div className="card">
            <div className="card-header btn" id="headingResumoProposta" data-toggle="collapse" data-target="#collapseResumoProposta" aria-expanded="true" aria-controls="collapseResumoProposta">
              <h5 className="mb-0 text-info" >
                Resumo da Proposta
              </h5>
            </div>
            <div id="collapseResumoProposta" className="collapse show" aria-labelledby="headingOne" >
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-8">
                      <label htmlFor="txtTitulo">Tipo de análise</label><br></br>
                      <span className="text-info" id='txtTipoAnalise'></span>
                    </div>
                  </div>
                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-9">
                        <label htmlFor="txtSintese">Síntese</label><br></br>
                        <span className="text-info" id='txtSintese'></span>
                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtIdentificacaoOportunidade">Identificação da Oportunidade </label><br></br>
                        <span className="text-info" id='txtIdentificacaoOportunidade'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataEntregaPropostaCliente">Data da entrega da Proposta ao Cliente</label><br></br>
                        <span className="text-info" id='txtDataEntregaPropostaCliente'></span>
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataFinalQuestionamentos">Data final de questionamentos</label><br></br>
                        <span className="text-info" id='txtDataFinalQuestionamentos'></span>
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataValidadeProposta">Data de validade da Proposta</label><br></br>
                        <span className="text-info" id='txtdataValidadeProposta'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlRepresentante">Representante</label><br></br>
                        <span className="text-info" id='txtRepresentante'></span>
                      </div>
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlCliente">Cliente </label><br></br>
                        <span className="text-info" id='txtCliente'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-8">
                        <label htmlFor="txtPropostaRevisadaReferencia">Proposta revisada/referência</label><br></br>
                        <span className="text-info" id='txtPropostaRevisadaReferencia'></span>
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtCondicoesPagamento">Condições de pagamento </label><br></br>
                        <span className="text-info" id='txtCondicoesPagamento'></span>
                      </div>
                    </div>
                  </div>

                </div>
              </div>
            </div>

          </div>

          <div className="card">
            <div className="card-header btn" id="headingDescricaoDetalhada" data-toggle="collapse" data-target="#collapseDescricaoDetalhada" aria-expanded="true" aria-controls="collapseDescricaoDetalhada">
              <h5 className="mb-0 text-info" >
                Descrição Detalhada
              </h5>
            </div>
            <div id="collapseDescricaoDetalhada" className="collapse show" aria-labelledby="headingOne" >
              <div className="card-body">

                <div className="form-group">
                  <label htmlFor="txtDadosProposta">Dados da Proposta</label><span className="required"> *</span>
                  <span id='txtDadosProposta'></span>
                </div>

              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingOportunidade" data-toggle="collapse" data-target="#collapseOportunidade" aria-expanded="true" aria-controls="collapseOportunidade">
              <h5 className="mb-0 text-info" >
                Oportunidade
              </h5>
            </div>
            <div id="collapseOportunidade" className="collapse show" aria-labelledby="headingOne" >
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-4">
                      <label htmlFor="txtSegmento">Segmento</label><br></br>
                      <span className="text-info" id='txtSegmento'></span>
                    </div>
                    <div className="form-group col-md-4">
                      <label htmlFor="txtSetor">Setor</label><br></br>
                      <span className="text-info" id='txtSetor'></span>
                    </div>
                    <div className="form-group col-md-4">
                      <label htmlFor="txtModalidade">Modalidade </label><br></br>
                      <span className="text-info" id='txtModalidade'></span>
                    </div>
                  </div>
                </div>
                <div className="form-group">
                  <label htmlFor="txtNumeroEditalRFPRFQRFI">Número do Edital, RFP, RFQ ou RFI </label><br></br>
                  <span className="text-info" id='txtNumeroEditalRFPRFQRFI'></span>
                </div>

              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingProduto" data-toggle="collapse" data-target="#collapseProduto" aria-expanded="true" aria-controls="collapseProduto">
              <h5 className="mb-0 text-info" >
                Produto
              </h5>
            </div>
            <div id="collapseProduto" className="collapse show" aria-labelledby="headingOne" >
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-3">
                      <label htmlFor="txtQuantidade">Quantidade</label><br></br>
                      <span className="text-info" id='txtQuantidade'></span>
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="txtInstalacao">Instalação</label><br></br>
                      <span className="text-info" id='txtInstalacao'></span>
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="txtGarantia">Garantia</label><br></br>
                      <span className="text-info" id='txtGarantia'></span>
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="txtTitulo">Tipo de garantia </label><br></br>
                      <span className="text-info" id='txtTipoGarantia'></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-2">
                      <label htmlFor="txtTitulo">Prazo de garantia </label><br></br>
                      <span className="text-info" id='txtPrazoGarantia'></span>
                    </div>
                    <div className="form-group col-md-2">
                      <label htmlFor="txtOutrosServicos">Outros serviços</label><br></br>
                      <span className="text-info" id='txtOutrosServicos'></span>
                    </div>
                    <div className="form-group col-md-8">
                      <label htmlFor="ddlProduto">Produto</label><br></br>
                      <span className="text-info" id='txtProduto'></span>
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingArea" data-toggle="collapse" data-target="#collapseArea" aria-expanded="true" aria-controls="collapseArea">
              <h5 className="mb-0 text-info" >
                Áreas Responsáveis pela Proposta
              </h5>
            </div>
            <div id="collapseArea" className="collapse show" aria-labelledby="headingOne" >

              <div className="card-body">

                <label htmlFor="txtAreas">Áreas</label><br></br>
                <span className="text-info" id='txtAreas'></span>
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingAnexos" data-toggle="collapse" data-target="#collapseAnexos" aria-expanded="true" aria-controls="collapseAnexos">
              <h5 className="mb-0 text-info" >
                Anexos
              </h5>
            </div>
            <div id="collapseAnexos" className="collapse show" aria-labelledby="headingOne" >

              <div className="card-body">

                <div id='conteudoAnexo'></div>

              </div>
            </div>
          </div>

          <br></br>

          <div className="text-right">
            <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
          </div>


        </div>

      </div>

    );
  }



  protected getProposta() {

    console.log("entrou no proposta");

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('PropostasSAP')/items?$select=ID,Title,TipoAnalise,IdentificacaoOportunidade,DataEntregaPropostaCliente,DataFinalQuestionamentos,DataValidadeProposta,Representante/ID,Representante/Title,Cliente/ID,Cliente/Title,PropostaRevisadaReferencia,CondicoesPagamento,DadosProposta,Segmento,Setor,Modalidade,NumeroEditalRFPRFQRFI,Instalacao,Quantidade,Garantia,TipoGarantia,PrazoGarantia,OutrosServicos,Produto/ID,Produto/Title&$expand=Representante,Cliente,Produto&$filter=ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var tipoAnalise = resultData.d.results[i].TipoAnalise;
            var sintese = resultData.d.results[i].Title;
            var identificacaoOportunidade = resultData.d.results[i].IdentificacaoOportunidade;
            var dataEntregaPropostaCliente = new Date(resultData.d.results[i].DataEntregaPropostaCliente);
            var dtDataEntregaPropostaCliente = ("0" + dataEntregaPropostaCliente.getDate()).slice(-2) + '/' + ("0" + (dataEntregaPropostaCliente.getMonth() + 1)).slice(-2) + '/' + dataEntregaPropostaCliente.getFullYear();
            var dataFinalQuestionamentos = new Date(resultData.d.results[i].DataFinalQuestionamentos);
            var dtdataFinalQuestionamentos = ("0" + dataFinalQuestionamentos.getDate()).slice(-2) + '/' + ("0" + (dataFinalQuestionamentos.getMonth() + 1)).slice(-2) + '/' + dataFinalQuestionamentos.getFullYear();
            var dataValidadeProposta = new Date(resultData.d.results[i].DataValidadeProposta);
            var dtdataValidadeProposta = ("0" + dataValidadeProposta.getDate()).slice(-2) + '/' + ("0" + (dataValidadeProposta.getMonth() + 1)).slice(-2) + '/' + dataValidadeProposta.getFullYear();
            var representante = resultData.d.results[i].Representante.Title;
            var cliente = resultData.d.results[i].Cliente.Title;
            var propostaRevisadaReferencia = resultData.d.results[i].PropostaRevisadaReferencia;
            var condicoesPagamento = resultData.d.results[i].CondicoesPagamento;
            var dadosProposta = resultData.d.results[i].DadosProposta;
            var arrSegmento = resultData.d.results[i].Segmento.results;
            var strSegmento = arrSegmento.toString();
            var arrSetor = resultData.d.results[i].Setor;
            var strSetor = arrSetor.toString();
            var arrModalidade = resultData.d.results[i].Modalidade;
            var strModalidade = arrModalidade.toString();
            var numeroEditalRFPRFQRFI = resultData.d.results[i].NumeroEditalRFPRFQRFI;
            var quantidade = resultData.d.results[i].Quantidade;
            var instalacao = resultData.d.results[i].Instalacao;
            var garantia = resultData.d.results[i].Garantia;
            var tipoGarantia = resultData.d.results[i].TipoGarantia;

            var prazoGarantia = resultData.d.results[i].PrazoGarantia;

            var arrOutrosServicos = resultData.d.results[i].OutrosServicos.results;
            var strOutrosServicos = arrOutrosServicos.toString();


            var arrProduto = resultData.d.results[i].Produto.results;
            var arrTituloProduto = [];

            for (i = 0; i < arrProduto.length; i++) {
              arrTituloProduto.push(arrProduto[i].Title);
            }

            var strTituloProduto = arrTituloProduto.toString();

            jQuery("#txtTipoAnalise").html(tipoAnalise);
            jQuery("#txtSintese").html(sintese);
            jQuery("#txtIdentificacaoOportunidade").html(identificacaoOportunidade);
            jQuery("#txtDataEntregaPropostaCliente").html(dtDataEntregaPropostaCliente);
            jQuery("#txtDataFinalQuestionamentos").html(dtdataFinalQuestionamentos);
            jQuery("#txtdataValidadeProposta").html(dtdataValidadeProposta);
            jQuery("#txtRepresentante").html(representante);
            jQuery("#txtCliente").html(cliente);
            jQuery("#txtPropostaRevisadaReferencia").html(propostaRevisadaReferencia);
            jQuery("#txtCondicoesPagamento").html(condicoesPagamento);
            jQuery("#txtDadosProposta").html(dadosProposta);
            jQuery("#txtSegmento").html(strSegmento);
            jQuery("#txtSetor").html(strSetor);
            jQuery("#txtModalidade").html(strModalidade);
            jQuery("#txtNumeroEditalRFPRFQRFI").html(numeroEditalRFPRFQRFI);
            jQuery("#txtQuantidade").html(quantidade);
            jQuery("#txtInstalacao").html(instalacao);
            jQuery("#txtGarantia").html(garantia);
            jQuery("#txtTipoGarantia").html(tipoGarantia);
            jQuery("#txtPrazoGarantia").html(prazoGarantia);
            jQuery("#txtOutrosServicos").html(strOutrosServicos);
            jQuery("#txtProduto").html(strTituloProduto);

          }

        }



      },
      error: function (jqXHR, textStatus, errorThrown) {
      }



    })

  }


  protected getTarefas() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        var arrAreas = [];

        console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            arrAreas.push(resultData.d.results[i].GrupoSharepoint.Title);

          }

        }

        jQuery("#txtAreas").html(arrAreas.toString());


      },
      error: function (jqXHR, textStatus, errorThrown) {
      }



    })


  }

  protected getAnexos() {

    //get anexos da biblioteca

    var montaAnexo = "";

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Proposta-Detalhes.aspx", "");

    //var relative = "/sites/bit-hml";
    var idItem = 0;

    console.log("_bitNumero", _idProposta);

    console.log("caminho", `${strRelativeURL}/AnexosSAP/${_idProposta}`);


    _web.getFolderByServerRelativeUrl(`${strRelativeURL}/AnexosSAP/${_idProposta}`)
      .expand("Folders, Files, ListItemAllFields").get().then(r => {
        console.log("r", r);
        /*
        r.Folders.forEach(item => {
          console.log("item-doc", item);
          console.log("entrou em folder");
        })
        */
        r.Files.forEach(item => {
          console.log("entrou em files");

          console.log("item", item);
          idItem++;
          $("#conteudoAnexoNaoEncontrado").hide();
          montaAnexo = `<a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a>&nbsp; <br/>`

          $("#conteudoAnexo").append(montaAnexo);


        })

      }).catch((error: any) => {
        console.log("Erro Anexo da biblioteca: ", error);
      });


    //fim anexos da biblioteca


  }
}
