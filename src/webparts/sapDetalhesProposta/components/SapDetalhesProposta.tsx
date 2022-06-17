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
import { first, replace } from 'lodash';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");
require("../../../../css/discussao.css");


var _idProposta;
var _web;
var _mensagemDiscussao;
var _siteURL;
var _idTarefa;
var _numeroProposta;
var _grupos = [];
var _strGrupos;
var _testeGus;
var _threadDiscussao;
var _caminho;
var _siteAntigo;
var _pastaCriada;
var _idAntigo;
var _arrAreas = [];
var _montaNotificarArea = "";
var _arrMontaBotoes = ["Todos"];

export interface IReactGetItemsState {
  itemsTarefas: [
    {
      "ID": "",
      "Title": "",
      "Status": string,
      "DataPlanejadaTermino": "",
      "DataRealTermino": "",
      "Justificativa": "",
      "GrupoSharepoint": { "Title": "" }
    }],
  itemsTarefasAntigas: [
    {
      "ID": "",
      "Title": any,
      "Status": string,
      "DataPlanejadaTermino": "",
      "DataRealTermino": "",
      "Justificativa": "",
      "GrupoSharepoint": { "Title": "" }
    }],

  itemAreas: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsBotoes: any
}

export default class SapDetalhesProposta extends React.Component<ISapDetalhesPropostaProps, IReactGetItemsState> {


  public constructor(props: ISapDetalhesPropostaProps, state: IReactGetItemsState) {
    super(props);
    this.state = {
      itemsTarefas: [
        {
          "ID": "",
          "Title": "",
          "Status": "",
          "DataPlanejadaTermino": "",
          "DataRealTermino": "",
          "Justificativa": "",
          "GrupoSharepoint": { "Title": "" }
        }],
      itemsTarefasAntigas: [
        {
          "ID": "",
          "Title": "",
          "Status": "",
          "DataPlanejadaTermino": "",
          "DataRealTermino": "",
          "Justificativa": "",
          "GrupoSharepoint": { "Title": "" }
        }],
      itemAreas: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsBotoes: []

    };
  }


  public async componentDidMount() {

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idProposta = parseInt(queryParms.getValue("PropostasID"));

    $('[aria-describedby="image-richtextbutton"]').hide();

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _siteURL = this.props.siteurl;

    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    jQuery("#btnEditarProposta").hide();
    jQuery("#conteudoUpload").hide();

    document
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());

    //let groups = await _web.currentUser.groups();
    await _web.currentUser.get().then(f => {
      // console.log("user", f);
      var id = f.Id;

      var grupos = [];

      jQuery.ajax({
        url: `${this.props.siteurl}/_api/web/GetUserById(${id})/Groups`,
        type: "GET",
        headers: { 'Accept': 'application/json; odata=verbose;' },
        async: false,
        success: async function (resultData) {

          //console.log("resultDataGrupo", resultData);

          if (resultData.d.results.length > 0) {

            for (var i = 0; i < resultData.d.results.length; i++) {

              grupos.push(resultData.d.results[i].Title);

            }

          }

        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(jqXHR.responseText);
        }

      })

      //console.log("grupos", grupos);
      _grupos = grupos;

    })



    document
      .getElementById("btnResponderDiscussao")
      .addEventListener("click", (e: Event) => this.modalResponderDiscussao());

    document
      .getElementById("btnCadastrarDiscussao")
      .addEventListener("click", (e: Event) => this.cadastrarDiscussao());

    document
      .getElementById("btnAprovarTarefa")
      .addEventListener("click", (e: Event) => this.aprovar());

    document
      .getElementById("btnReabrirProposta")
      .addEventListener("click", (e: Event) => this.modalReabrirProposta());

    document
      .getElementById("btnModalReabrirProposta")
      .addEventListener("click", (e: Event) => this.reabrirProposta());

    document
      .getElementById("btnEditarProposta")
      .addEventListener("click", (e: Event) => this.editarProposta());

    document
      .getElementById("btnSucesso")
      .addEventListener("click", (e: Event) => this.fecharSucesso());

    document
      .getElementById("btnAnexar")
      .addEventListener("click", (e: Event) => this.upload());


    $("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/Images1/loading.gif'/>
    <br/>Aguarde....<br/><br/>
    Dependendo do tamanho do anexo e a velocidade<br>
     da Internet essa ação pode demorar um pouco. <br>
     Não fechar a janela!<br/><br/>`);

    this.getProposta();
    this.getTarefas();
    this.getAnexos();
    this.getSelecaoAreas();
    this.getDiscussao();

  }



  public render(): React.ReactElement<ISapDetalhesPropostaProps> {
    return (

      <>
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
                    <div className="form-row ">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtNumeroProposta">Número da Proposta</label><br></br>
                        <span className="text-info" id='txtNumeroProposta'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtStatus">Status</label><br></br>
                        <span className="text-info" id='txtStatus'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtTipoAnalise">Tipo de análise</label><br></br>
                        <span className="text-info" id='txtTipoAnalise'></span>
                      </div>
                    </div>
                  </div>




                  <div className="form-group">
                    <div className="form-row ">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtSintese">Síntese</label><br></br>
                        <span className="text-info" id='txtSintese'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtIdentificacaoOportunidade">Identificação da Oportunidade </label><br></br>
                        <span className="text-info" id='txtIdentificacaoOportunidade'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row ">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="dtDataEntregaPropostaCliente">Data da entrega da Proposta ao Cliente</label><br></br>
                        <span className="text-info" id='txtDataEntregaPropostaCliente'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="dtDataFinalQuestionamentos">Data final de questionamentos</label><br></br>
                        <span className="text-info" id='txtDataFinalQuestionamentos'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="dtDataValidadeProposta">Data de validade da Proposta</label><br></br>
                        <span className="text-info" id='txtdataValidadeProposta'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row ">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="ddlRepresentante">Representante</label><br></br>
                        <span className="text-info" id='txtRepresentante'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="ddlCliente">Cliente </label><br></br>
                        <span className="text-info" id='txtCliente'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row ">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtPropostaRevisadaReferencia">Responsável pela Proposta</label><br></br>
                        <span className="text-info" id='txtResponsavelPelaProposta'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtPropostaRevisadaReferencia">Proposta revisada/referência</label><br></br>
                        <span className="text-info" id='txtPropostaRevisadaReferencia'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtCondicoesPagamento">Condições de pagamento </label><br></br>
                        <span className="text-info" id='txtCondicoesPagamento'></span>
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
                    <div className="form-row ">
                      <div className="form-group col-md border m-1">
                        <label htmlFor="txtDadosProposta">Dados da Proposta</label><span className="required"> *</span>
                        <span id='txtDadosProposta'></span>
                      </div>
                    </div>
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
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtSegmento">Segmento</label><br></br>
                        <span className="text-info" id='txtSegmento'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtSetor">Setor</label><br></br>
                        <span className="text-info" id='txtSetor'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtModalidade">Modalidade </label><br></br>
                        <span className="text-info" id='txtModalidade'></span>
                      </div>
                    </div>
                  </div>
                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtNumeroEditalRFPRFQRFI">Número do Edital, RFP, RFQ ou RFI </label><br></br>
                        <span className="text-info" id='txtNumeroEditalRFPRFQRFI'></span>
                      </div>
                    </div>
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
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtQuantidade">Quantidade</label><br></br>
                        <span className="text-info" id='txtQuantidade'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtInstalacao">Instalação</label><br></br>
                        <span className="text-info" id='txtInstalacao'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtGarantia">Garantia</label><br></br>
                        <span className="text-info" id='txtGarantia'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtTitulo">Tipo de garantia </label><br></br>
                        <span className="text-info" id='txtTipoGarantia'></span>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtTitulo">Prazo de garantia </label><br></br>
                        <span className="text-info" id='txtPrazoGarantia'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
                        <label htmlFor="txtOutrosServicos">Outros serviços</label><br></br>
                        <span className="text-info" id='txtOutrosServicos'></span>
                      </div>
                      <div className="form-group col-md border m-1" style={{ "height": "53px" }}>
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

                  <div className="form-group">
                    <div className="form-row ">
                      <div className="form-group col-md border m-1">
                        <label htmlFor="txtAreas">Áreas</label><br></br>
                        <span className="text-info" id='txtAreas'></span>
                      </div>
                    </div>
                  </div>

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

                  <div className="table-responsive">

                    <table className="table table-hover" id="tbItens">
                      <thead>
                        <tr>
                          <th scope="col">Nome</th>
                          <th scope="col">Área</th>
                          <th scope="col">Criado</th>
                          <th scope="col">Criado por</th>
                        </tr>
                      </thead>
                      <tbody id="conteudoAnexo">
                      </tbody>
                    </table>

                  </div>

                  <br></br>
                  <div id='conteudoUpload' className='form-group col-md border m-1'>
                    <br></br>
                    <div className="form-group">
                      <div className="form-row ">
                        <div className="form-group col-md" >
                          <label htmlFor="input">Anexo </label><span className="required"> *</span>
                          <input className="multi" data-maxsize="1024" type="file" id="input" multiple />
                        </div>
                        <div className="form-group col-md" >
                          <label htmlFor="ddlAreaAnexo">Área </label><span className="required"> *</span>
                          <select id="ddlAreaAnexo" className="form-control" style={{ "width": "300px" }}>
                            <option value="0" selected>Selecione...</option>
                            {this.state.itemAreas.map(function (item, key) {
                              return (
                                <option value={item.ID}>{item.Title}</option>
                              );
                            })}
                          </select>
                        </div>
                        <div className="form-group col-md" >
                          <label htmlFor="btnAnexar">&nbsp;</label><br></br>
                          <button id='btnAnexar' type="button" className="btn btn-secondary ">Anexar</button>
                        </div>
                      </div>
                    </div>
                    <p className='text-info'>Total máximo permitido: 15 MB</p>

                  </div>


                </div>
              </div>
            </div>

            <div className="card">
              <div className="card-header btn" id="headingDiscussao" data-toggle="collapse" data-target="#collapseDiscussao" aria-expanded="true" aria-controls="collapseDiscussao">
                <h5 className="mb-0 text-info" >
                  Discussão
                </h5>
              </div>
              <div id="collapseDiscussao" className="collapse show" aria-labelledby="headingOne" >

                <div className="card-body">

                  <div id="conteudo_botao"></div>

                  {this.state.itemsBotoes.map((item, key) => {

                    var area2 = item.replaceAll(/[^A-Z0-9]+/ig, "_");

                    return (

                      <button id="bt_${idBtn}" onClick={(e) => this.filtrar(area2)} type="button" className="btn btn-light btn-sm ml-1">{item}</button>

                    );
                  })}

                  <br></br><br></br>

                  <div id='conteudoDiscusssaoAntiga'></div>
                  <div id='conteudoDiscusssaoNova'></div>

                  <br></br>
                  <button id='btnResponderDiscussao' type="button" className="btn btn-info">Responder</button>
                  <br></br>


                </div>

              </div>
            </div>

            <div className="card">
              <div className="card-header btn" id="headingTarefas" data-toggle="collapse" data-target="#collapseTarefas" aria-expanded="true" aria-controls="collapseTarefas">
                <h5 className="mb-0 text-info" >
                  Aprovação
                </h5>
              </div>
              <div id="collapseTarefas" className="collapse show" aria-labelledby="headingOne" >

                <div className="card-body">

                  <table className="table table-hover" id="tbItens">
                    <thead>
                      <tr>
                        <th scope="col">Avaliação da Área</th>
                        <th scope="col">Status</th>
                        <th scope="col">Data planejada</th>
                        <th scope="col">Data real</th>
                        <th scope="col">Justificativa</th>
                        <th scope="col">Ação</th>
                      </tr>
                    </thead>
                    <tbody id="conteudoTarefas">
                      {this.state.itemsTarefas.map((item) => {

                        var dataPlanejadaTermino = new Date(item.DataPlanejadaTermino);
                        var dtDataPlanejadaTermino = ("0" + dataPlanejadaTermino.getDate()).slice(-2) + '/' + ("0" + (dataPlanejadaTermino.getMonth() + 1)).slice(-2) + '/' + dataPlanejadaTermino.getFullYear();
                        var dataRealTermino = new Date(item.DataRealTermino);
                        var dtDataRealTermino = ("0" + dataRealTermino.getDate()).slice(-2) + '/' + ("0" + (dataRealTermino.getMonth() + 1)).slice(-2) + '/' + dataRealTermino.getFullYear();
                        if (dtDataRealTermino == "31/12/1969") dtDataRealTermino = "-";
                        var vlrJustificativa;
                        var justificativa = item.Justificativa;

                        if (justificativa == null) vlrJustificativa = "-";
                        else vlrJustificativa = justificativa;

                        // var status = item.Status;

                        //console.log("item.Status", item.Status);

                        if (item.Status == "Em análise") {

                          //console.log("item.Title", item.Title);
                          //console.log("_grupos.indexOf(item.Title)", _grupos.indexOf(item.Title));

                          if (_grupos.indexOf(item.Title) !== -1) {

                            jQuery("#conteudoUpload").show();

                            return (

                              <><tr>
                                <td className='break-word'>{item.Title}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td className='break-word'>{vlrJustificativa}</td>
                                <td><button onClick={() => this.detalhesTarefas(item.ID, item.GrupoSharepoint.Title)} type="button" className="btn btn-success btn-sm">Detalhes</button></td>
                              </tr></>
                            );

                          } else {

                            return (

                              <><tr>
                                <td className='break-word'>{item.Title}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td className='break-word'>{vlrJustificativa}</td>
                                <td className='break-word'></td>
                              </tr></>
                            );

                          }


                        } else {

                          return (

                            <><tr>
                              <td className='break-word'>{item.Title}</td>
                              <td>{item.Status}</td>
                              <td>{dtDataPlanejadaTermino}</td>
                              <td>{dtDataRealTermino}</td>
                              <td className='break-word'>{vlrJustificativa}</td>
                              <td className='break-word'></td>
                            </tr></>
                          );


                        }



                      })}
                      {this.state.itemsTarefasAntigas.map((item) => {

                        var titulo = item.Title;
                        var rplTitulo = titulo.replaceAll("Avaliação da Área (", "");
                        var rplTitulo = rplTitulo.replaceAll(")", "");

                        var dataPlanejadaTermino = new Date(item.DataPlanejadaTermino);
                        var dtDataPlanejadaTermino = ("0" + dataPlanejadaTermino.getDate()).slice(-2) + '/' + ("0" + (dataPlanejadaTermino.getMonth() + 1)).slice(-2) + '/' + dataPlanejadaTermino.getFullYear();
                        var dataRealTermino = new Date(item.DataRealTermino);
                        var dtDataRealTermino = ("0" + dataRealTermino.getDate()).slice(-2) + '/' + ("0" + (dataRealTermino.getMonth() + 1)).slice(-2) + '/' + dataRealTermino.getFullYear();
                        if (dtDataRealTermino == "31/12/1969") dtDataRealTermino = "-";
                        var vlrJustificativa;
                        var justificativa = item.Justificativa;

                        if (justificativa == null) vlrJustificativa = "-";
                        else vlrJustificativa = justificativa;

                        // var status = item.Status;

                        //console.log("item.Status", item.Status);

                        if (item.Status == "Em análise") {

                          //console.log("item.Title", item.Title);
                          //console.log("_grupos.indexOf(item.Title)", _grupos.indexOf(item.Title));

                          if (_grupos.indexOf(item.Title) !== -1) {

                            jQuery("#conteudoUpload").show();

                            return (

                              <><tr>
                                <td className='break-word'>{rplTitulo}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td className='break-word'>{vlrJustificativa}</td>
                                <td className='break-word'><button onClick={() => this.detalhesTarefas(item.ID, item.GrupoSharepoint.Title)} type="button" className="btn btn-success btn-sm">Detalhes</button></td>
                              </tr></>
                            );

                          } else {

                            return (

                              <><tr>
                                <td className='break-word'>{rplTitulo}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td className='break-word'>{vlrJustificativa}</td>
                                <td className='break-word'></td>
                              </tr></>
                            );

                          }


                        } else {

                          return (

                            <><tr>
                              <td className='break-word'>{rplTitulo}</td>
                              <td>{item.Status}</td>
                              <td>{dtDataPlanejadaTermino}</td>
                              <td>{dtDataRealTermino}</td>
                              <td className='break-word'>{vlrJustificativa}</td>
                              <td className='break-word'></td>
                            </tr></>
                          );


                        }



                      })}

                    </tbody>
                  </table>

                </div>

              </div>
            </div>

            <br></br>

            <span className='text-info'>Criador por: <span id='txtCriadoPor'></span> no dia <span id='txtCriadoData'></span> às <span id='txtCriadoHora'>11:30</span></span><br></br>
            <span className='text-info'>Modificado por: <span id='txtModificadoPor'></span> no dia <span id='txtModificadoData'></span> às <span id='txtModificadoHora'>11:30</span></span>


            <br></br>



            <div className="text-right">
              <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
              <button style={{ "margin": "2px" }} type="submit" id="btnReabrirProposta" className="btn btn-danger">Reabrir Proposta</button>
              <button style={{ "margin": "2px" }} type="submit" id="btnEditarProposta" className="btn btn-primary">Editar</button>

            </div>


          </div>

        </div>

        <div className="modal fade" id="modalCarregando" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div>
            <div className="modal-dialog" role="document">
              <div className="modal-content">
                <div id='conteudoLoading' className='carregando'></div>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalDiscussao" tabIndex={-1} role="dialog" aria-labelledby="discussaoTitle" aria-hidden="true">
          <div className="modal-dialog modal-dialog-scrollable" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="discussaoTitle">Cadastrar Discussão</h5>
              </div>
              <div className="modal-body">
                <div className="container-fluid border rounded table-responsive" id="conteudo_formulario">
                  <div className="form-group">
                    <label htmlFor="txtDescricao">Área</label><span className={styles.required}> *</span>
                    <select id="ddlArea" className="form-control">
                    </select>
                  </div>
                  <div className="form-group">
                    <label htmlFor="txtDescricao">Mensagem</label><span className={styles.required}> *</span>
                    <RichText value="" className="editorRichTex"
                      onChange={(text) => this.onTextChangeMensagemDiscussao(text)}
                    />
                  </div>
                  <div className="form-group">
                    <label htmlFor="txtDescricao">Anexos relacionados</label>
                    <div id='conteudoAnexosRelacionados'></div>
                  </div>
                  <div className="form-group">
                    <label htmlFor="txtDescricao">Notificar Área</label><span className={styles.required}> *</span>
                    <div id='conteudoNotificarArea'></div>

                  </div>
                </div>
              </div>


              <div className="modal-footer">
                <button type="button" id='btnCadastrarDiscussaoCancelar' className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button type="button" id='btnCadastrarDiscussao' className="btn btn-primary">Cadastrar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalAprovacao" tabIndex={-1} role="dialog" aria-labelledby="aprovacaoTitle" aria-hidden="true">
          <div className="modal-dialog modal-dialog-scrollable" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="aprovacaoTitle">Aprovar</h5>
              </div>
              <div className="modal-body">
                <div className="container-fluid border rounded table-responsive" id="conteudo_formulario">
                  <div className="form-group">
                    <h4>Avaliação da Área <span id='txtModalArea'></span></h4>
                    <span className="text-info">Proposta nro: <span id='txtModalNumeroProposta'></span></span><br></br><br></br>
                    <label htmlFor="ddlStatus">Status</label><span className={styles.required}> *</span>
                    <select id="ddlStatus" className="form-control">
                      <option value="0" selected>Selecione...</option>
                      <option value="Aprovada">Aprovada</option>
                      <option value="Reprovada">Reprovada</option>
                      <option value="Não envolve a Área">Não envolve a Área</option>
                    </select>
                  </div>
                  <div className="form-group">
                    <label htmlFor="txtDescricao">Justificativa</label><span className={styles.required}> *</span>
                    <textarea id="txtJustificativa" className="form-control" rows={4}></textarea>
                  </div>
                </div>
              </div>

              <div className="modal-footer">
                <button type="button" id='btnAprovarTarefaCancelar' className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button type="button" id='btnAprovarTarefa' className="btn btn-primary">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalReabrirProposta" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Reabertura da Proposta <span id='txtModalNumeroProposta'></span></h5>
              </div>
              <div className="modal-body">
                <div className="form-group">
                  <label htmlFor="ddlStatusReabertura">Status</label><span className={styles.required}> *</span>
                  <select id="ddlStatusReabertura" className="form-control">
                    <option value="0" selected>Selecione...</option>
                    <option value="Voltar para em andamento">Voltar para em andamento</option>
                    <option value="Vencedora">Vencedora</option>
                    <option value="Não Vencedora">Não Vencedora</option>
                    <option value="Reprovada">Reprovada</option>
                    <option value="Cancelada">Cancelada</option>
                  </select>
                </div>
                <div className="form-group">
                  <label htmlFor="ddlMotivoReabertura">Motivo</label><span className={styles.required}> *</span>
                  <select id="ddlMotivoReabertura" className="form-control">
                    <option value="0" selected>Selecione...</option>
                    <option value="Preço">Preço</option>
                    <option value="Solução integrada">Solução integrada</option>
                    <option value="Reprovada pelo cliente">Reprovada pelo cliente</option>
                    <option value="Reprovada pela Doebold">Reprovada pela Doebold</option>
                    <option value="Proposta cancelada">Proposta cancelada</option>
                    <option value="Nova revisão">Nova revisão</option>
                  </select>
                </div>
                <div className="form-group">
                  <label htmlFor="txtJustificativaReabertura">Justificativa</label><span className={styles.required}> *</span>
                  <textarea id="txtJustificativaReabertura" className="form-control" rows={4}></textarea>
                </div>

              </div>
              <div className="modal-footer">
                <button type="button" id='btnModalReabrirPropostaCancelar' className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnModalReabrirProposta" type="button" className="btn btn-primary">Reabrir Proposta</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucesso" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Arquivo anexado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucesso" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

      </>



    );
  }



  protected getProposta() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('PropostasSAP')/items?$select=ID,Title,TipoAnalise,IdentificacaoOportunidade,DataEntregaPropostaCliente,DataFinalQuestionamentos,DataValidadeProposta,Representante/ID,Representante/Title,Cliente/ID,Cliente/Title,PropostaRevisadaReferencia,CondicoesPagamento,DadosProposta,Segmento,Setor,Modalidade,NumeroEditalRFPRFQRFI,Instalacao,Quantidade,Garantia,TipoGarantia,PrazoGarantia,OutrosServicos,Produto/ID,Produto/Title,Numero,Status,ResponsavelProposta,Created,Author/Title,Modified,Editor/Title,SiteAntigo,PastaCriada,IDAntigo&$expand=Representante,Cliente,Produto,Author,Editor&$filter=ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        //console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            _siteAntigo = resultData.d.results[i].SiteAntigo;
            _pastaCriada = resultData.d.results[i].PastaCriada;
            _idAntigo = resultData.d.results[i].IDAntigo;

            var numeroProposta = resultData.d.results[i].Numero;
            _numeroProposta = numeroProposta;
            var status = resultData.d.results[i].Status;
            var tipoAnalise = resultData.d.results[i].TipoAnalise;
            var sintese = resultData.d.results[i].Title;
            var identificacaoOportunidade = resultData.d.results[i].IdentificacaoOportunidade;
            var dataEntregaPropostaCliente = new Date(resultData.d.results[i].DataEntregaPropostaCliente);
            var dtDataEntregaPropostaCliente = ("0" + dataEntregaPropostaCliente.getDate()).slice(-2) + '/' + ("0" + (dataEntregaPropostaCliente.getMonth() + 1)).slice(-2) + '/' + dataEntregaPropostaCliente.getFullYear();
            var dataFinalQuestionamentos = new Date(resultData.d.results[i].DataFinalQuestionamentos);
            var dtdataFinalQuestionamentos = ("0" + dataFinalQuestionamentos.getDate()).slice(-2) + '/' + ("0" + (dataFinalQuestionamentos.getMonth() + 1)).slice(-2) + '/' + dataFinalQuestionamentos.getFullYear();
            if (dtdataFinalQuestionamentos == "31/12/1969") dtdataFinalQuestionamentos = "";

            var dataValidadeProposta = new Date(resultData.d.results[i].DataValidadeProposta);
            var dtdataValidadeProposta = ("0" + dataValidadeProposta.getDate()).slice(-2) + '/' + ("0" + (dataValidadeProposta.getMonth() + 1)).slice(-2) + '/' + dataValidadeProposta.getFullYear();
            if (dtdataValidadeProposta == "31/12/1969") dtdataValidadeProposta = "";

            var representante = resultData.d.results[i].Representante.Title;
            var cliente = resultData.d.results[i].Cliente.Title;
            var responsavelProposta = resultData.d.results[i].ResponsavelProposta;
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

            var outrosServicos = resultData.d.results[i].OutrosServicos;

            if (outrosServicos != null) {
              var arrOutrosServicos = resultData.d.results[i].OutrosServicos.results;
              var strOutrosServicos = arrOutrosServicos.toString();
            }
            else strOutrosServicos = "-"


            var arrProduto = resultData.d.results[i].Produto.results;
            var arrTituloProduto = [];

            for (var x = 0; x < arrProduto.length; x++) {
              arrTituloProduto.push(arrProduto[x].Title);
            }

            var strTituloProduto = arrTituloProduto.toString();

            var criadoPor = resultData.d.results[i].Author.Title;
            var modificadoPor = resultData.d.results[i].Editor.Title;

            var criado = new Date(resultData.d.results[i].Created);
            var criadoData = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear();
            var criadoHora = criado.getHours() + ":" + ("0" + (criado.getMinutes() + 1)).slice(-2) + ":" + criado.getSeconds();

            var modificado = new Date(resultData.d.results[i].Modified);
            var modificadoData = ("0" + modificado.getDate()).slice(-2) + '/' + ("0" + (modificado.getMonth() + 1)).slice(-2) + '/' + modificado.getFullYear();;
            var modificadoHora = modificado.getHours() + ":" + ("0" + (modificado.getMinutes() + 1)).slice(-2) + ":" + modificado.getSeconds();

            jQuery("#txtCriadoPor").html(criadoPor);
            jQuery("#txtCriadoData").html(criadoData);
            jQuery("#txtCriadoHora").html(criadoHora);
            jQuery("#txtModificadoPor").html(modificadoPor);
            jQuery("#txtModificadoData").html(modificadoData);
            jQuery("#txtModificadoHora").html(modificadoHora);

            jQuery("#txtNumeroProposta").html(numeroProposta);
            jQuery("#txtStatus").html(status);
            jQuery("#txtTipoAnalise").html(tipoAnalise);
            jQuery("#txtSintese").html(sintese);
            jQuery("#txtIdentificacaoOportunidade").html(identificacaoOportunidade);
            jQuery("#txtDataEntregaPropostaCliente").html(dtDataEntregaPropostaCliente);
            jQuery("#txtDataFinalQuestionamentos").html(dtdataFinalQuestionamentos);
            jQuery("#txtdataValidadeProposta").html(dtdataValidadeProposta);
            jQuery("#txtRepresentante").html(representante);
            jQuery("#txtCliente").html(cliente);
            jQuery("#txtResponsavelPelaProposta").html(responsavelProposta);
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
            jQuery("#txtModalNumeroProposta").html(numeroProposta);

            console.log("status", status);

            if (status == "Em análise") {

              if (_grupos.indexOf("Representante") !== -1) {

                jQuery("#btnEditarProposta").show();

              }

              if (_grupos.indexOf("Comercial") !== -1) {

                jQuery("#btnEditarProposta").show();

              }

            }

            if (status != "Em análise") {

              jQuery("#btnResponderDiscussao").hide();

            }

            if (status != "Aprovado") {
              jQuery("#btnReabrirProposta").hide();
            }

          }

        }



      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }



    })

  }


  protected getTarefas() {


    var reactHandlerTarefasAntigas = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=AntigoPropostaNumero eq ` + _numeroProposta + `.000000000`,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        reactHandlerTarefasAntigas.setState({
          itemsTarefasAntigas: resultData.d.results
        });

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            _arrAreas.push(resultData.d.results[i].GrupoSharepoint.Title);

          }

        }

        for (var i = 0; i < _arrAreas.length; i++) {

          _montaNotificarArea += `<div class="form-check">
                                  <input class="form-check-input" name="checkNotificarArea" grupo="${resultData.d.results[i].GrupoSharepoint.Title}" type="checkbox" value="${resultData.d.results[i].GrupoSharepoint.ID}">
                                  <label class="form-check-label">${resultData.d.results[i].GrupoSharepoint.Title}</label>
                                  </div>`;

        }




      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });



    var reactHandlerTarefas = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        reactHandlerTarefas.setState({
          itemsTarefas: resultData.d.results,
          itemAreas: resultData.d.results
        });

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            _arrAreas.push(resultData.d.results[i].GrupoSharepoint.Title);

          }

        }

        for (var i = 0; i < _arrAreas.length; i++) {

          _montaNotificarArea += `<div class="form-check">
                                  <input class="form-check-input" name="checkNotificarArea" grupo="${resultData.d.results[i].GrupoSharepoint.Title}" type="checkbox" value="${resultData.d.results[i].GrupoSharepoint.ID}">
                                  <label class="form-check-label">${resultData.d.results[i].GrupoSharepoint.Title}</label>
                                  </div>`;

        }



      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    console.log("_arrAreas.toString()", _arrAreas.toString());
    jQuery("#txtAreas").html(_arrAreas.toString());
    $("#conteudoNotificarArea").append(_montaNotificarArea);


  }

  protected detalhesTarefas(id, nomeGrupo) {

    //console.log("entrou no detalhesTarefas" + id);
    _idTarefa = id;

    jQuery("#txtModalArea").html(nomeGrupo);
    jQuery("#modalAprovacao").modal({ backdrop: 'static', keyboard: false });
    $("#btnAprovarTarefa").prop("disabled", false);
    $("#btnAprovarTarefaCancelar").prop("disabled", false);

  }

  protected async aprovar() {

    $("#btnAprovarTarefa").prop("disabled", true);
    $("#btnAprovarTarefaCancelar").prop("disabled", true);


    var status = jQuery('#ddlStatus option:selected').val();
    var justificativa = jQuery('#txtJustificativa').val();

    if (status == "0") {
      alert("Escolha um Status");
      $("#btnAprovarTarefa").prop("disabled", false);
      $("#btnAprovarTarefaCancelar").prop("disabled", false);
      return false;
    }

    if (justificativa == "") {
      alert("Forneça uma justificativa");
      $("#btnAprovarTarefa").prop("disabled", false);
      $("#btnAprovarTarefaCancelar").prop("disabled", false);
      return false;
    }

    var DataRealTermino = "" + jQuery("#dtDataFinalQuestionamentos-label").val() + "";
    var DataRealTerminoDia = DataRealTermino.substring(0, 2);
    var DataRealTerminoMes = DataRealTermino.substring(3, 5);
    var DataRealTerminoAno = DataRealTermino.substring(6, 10);
    var formDataRealTermino = DataRealTerminoAno + "-" + DataRealTerminoMes + "-" + DataRealTerminoDia;

    var today = new Date();
    var dataRealTermino = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();

    await _web.lists
      .getByTitle("Tarefas")
      .items.getById(_idTarefa).update({
        Status: status,
        Justificativa: justificativa,
        DataRealTermino: dataRealTermino
      })
      .then(async response => {


        if (status == "Aprovada") {

          jQuery.ajax({
            url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,Status&$filter=(Proposta/ID eq ${_idProposta}) and (Status ne 'Aprovada' and Status ne 'Não envolve a Área')`,
            type: "GET",
            headers: { 'Accept': 'application/json; odata=verbose;' },
            async: false,
            success: async (resultData) => {

              if (resultData.d.results.length > 0) {

                this.recarregaTarefas();
                $("#modalAprovacao").modal('hide');

              } else {

                await _web.lists
                  .getByTitle("PropostasSAP")
                  .items.getById(_idProposta).update({
                    Status: "Aprovado",
                  })
                  .then(response => {
                    this.recarregaTarefas();
                    $("#modalAprovacao").modal('hide');
                    window.location.href = `Proposta-Detalhes.aspx?PropostasID=${_idProposta}`;
                  }).catch((error: any) => {
                    console.log(error);
                  });

              }


            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(jqXHR.responseText);
            }

          })
        }


        else if (status == "Não envolve a Área") {

          jQuery.ajax({
            url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,Status&$filter=(Proposta/ID eq ${_idProposta}) and (Status ne 'Aprovada' and Status ne 'Não envolve a Área')`,
            type: "GET",
            headers: { 'Accept': 'application/json; odata=verbose;' },
            async: false,
            success: async (resultData) => {

              if (resultData.d.results.length > 0) {

                this.recarregaTarefas();
                $("#modalAprovacao").modal('hide');

              } else {

                await _web.lists
                  .getByTitle("PropostasSAP")
                  .items.getById(_idProposta).update({
                    Status: "Aprovado",
                  })
                  .then(response => {
                    this.recarregaTarefas();
                    $("#modalAprovacao").modal('hide');
                    window.location.href = `Proposta-Detalhes.aspx?PropostasID=${_idProposta}`;
                  }).catch((error: any) => {
                    console.log(error);
                  });

              }


            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(jqXHR.responseText);
            }

          })


        }

        else if (status == "Reprovada") {

          await _web.lists
            .getByTitle("PropostasSAP")
            .items.getById(_idProposta).update({
              Status: "Reprovado",

            }).then(response => {

              jQuery.ajax({
                url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title&$filter=Proposta/ID eq ${_idProposta}`,
                type: "GET",
                headers: { 'Accept': 'application/json; odata=verbose;' },
                async: false,
                success: async (resultData) => {

                  if (resultData.d.results.length > 0) {

                    for (var i = 0; i < resultData.d.results.length; i++) {

                      var idTarefa = resultData.d.results[i].ID;

                      await _web.lists
                        .getByTitle("Tarefas")
                        .items.getById(idTarefa).update({
                          Status: "Reprovada",

                        }).then(response => {

                          console.log("Reprovou a tarefa!!!");

                        }).catch((error: any) => {
                          console.log(error);
                        });

                    }
                    console.log("Reprovou!!!");
                    this.recarregaTarefas();
                    window.location.href = `Proposta-Detalhes.aspx?PropostasID=${_idProposta}`;
                  }

                },
                error: function (jqXHR, textStatus, errorThrown) {
                  console.log(jqXHR.responseText);
                }

              })

            }).catch((error: any) => {
              console.log(error);
            });

        }

      }).catch((error: any) => {
        console.log(error);
      });

  }

  protected getAnexos() {

    var montaAnexo = "";
    var montaAnexo2 = "";
    var montaAnexosRelacionados = "";
    var relativeURL = window.location.pathname;
    var strRelativeURL = relativeURL.replace("SitePages/Proposta-Detalhes.aspx", "");
    var idItem = 0;

    if (_siteAntigo == "Sim") {

      jquery.ajax({
        url: `${this.props.siteurl}/_api/web/lists/getbytitle('Anexos')/items?$select=Title,AreaSelecionada,Created,Author/Title,File/ServerRelativeUrl&$expand=Author,File&$filter=Proposta/Numero eq ` + _numeroProposta,
        type: "GET",
        async: false,
        headers: { 'Accept': 'application/json; odata=verbose;' },
        success: function (resultData) {

          //console.log("resultData anexos antigos", resultData);

          if (resultData.d.results.length > 0) {

            for (var i = 0; i < resultData.d.results.length; i++) {

              var criado = new Date(resultData.d.results[i].Created);
              var criadoData = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear();
              var criadoHora = criado.getHours() + ":" + ("0" + (criado.getMinutes() + 1)).slice(-2) + ":" + criado.getSeconds();

              montaAnexo2 += `<tr> 
              <td style="word-break: break-word;"><a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${resultData.d.results[i].File.ServerRelativeUrl}">${resultData.d.results[i].Title}</a></td>
              <td style="word-break: break-word;">${resultData.d.results[i].AreaSelecionada}</td>
              <td style="word-break: break-word;">${criadoData} ${criadoHora}</td>
              <td style="word-break: break-word;">${resultData.d.results[i].Author.Title}</td>
              </tr>
              `

            }

            //console.log("montaAnexo2",montaAnexo2);
            jQuery("#conteudoAnexo").append(montaAnexo2);

          }


        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(jqXHR.responseText);
        }
      });

    }

    if (_pastaCriada == "Sim") {

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
            //console.log("entrou em files");
            //console.log("item", item);
            idItem++;
            jQuery("#conteudoAnexoNaoEncontrado").hide();

            montaAnexosRelacionados = `<div class="form-check">
                                    <input class="form-check-input" name="checkAnexosSelecionados" type="checkbox" value="${item.Name}">
                                    <label class="form-check-label">${item.Name}</label>
                                    </div>`

            jQuery("#conteudoAnexosRelacionados").append(montaAnexosRelacionados);


          })

        }).catch((error: any) => {
          console.log("Erro Anexo da biblioteca: ", error);
        });




      _web.getFolderByServerRelativeUrl(`${strRelativeURL}/AnexosSAP/${_idProposta}`).files
        .expand('ListItemAllFields', 'Author').get().then(r => {

          console.log("r1", r);

          if (r.length != 0) {

            r.forEach(item => {

              var criado = new Date(item.TimeCreated);
              var criadoData = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear();
              var criadoHora = criado.getHours() + ":" + ("0" + (criado.getMinutes() + 1)).slice(-2) + ":" + criado.getSeconds();

              montaAnexo += `<tr> 
            <td style="word-break: break-word;"><a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a></td>
            <td style="word-break: break-word;">${item.ListItemAllFields.Area}</td>
            <td style="word-break: break-word;">${criadoData} ${criadoHora}</td>
            <td style="word-break: break-word;">${item.Author.Title}</td>
            </tr>
            `

            });

            jQuery("#conteudoAnexo").append(montaAnexo);

          }

        }).catch((error: any) => {
          console.log("Erro onChangeCliente: ", error);
        });

    }



  }

  protected getSelecaoAreas() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Areas')/items?$select=ID,Title,AreaGrupo/ID,AreaGrupo/Title&$expand=AreaGrupo&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        var montaArea = `<option value="0" selected>Selecione...</option>`;

        // console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            if (_grupos.indexOf(resultData.d.results[i].AreaGrupo.Title) !== -1) {

              montaArea += `<option value=${resultData.d.results[i].ID}>${resultData.d.results[i].AreaGrupo.Title}</option>`;

            }

          }

        }

        $("#ddlArea").append(montaArea);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }

    })


  }

  protected getSelecaoAreasOLD() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Seleção da Área')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title&$expand=GrupoSharepoint&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        var montaArea = `<option value="0" selected>Selecione...</option>`;

        // console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            montaArea += `<option value=${resultData.d.results[i].ID}>${resultData.d.results[i].GrupoSharepoint.Title}</option>`;

          }

        }

        $("#ddlArea").append(montaArea);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }

    })


  }

  protected async getDiscussao() {

    if (_idAntigo != null) {

      jQuery.ajax({
        url: `${this.props.siteurl}/_api/web/lists/getbytitle('DiscussaoAntiga')/items?$top=1&$select=ID,Title,Thread&$filter=Title eq ` + _idAntigo,
        type: "GET",
        headers: { 'Accept': 'application/json; odata=verbose;' },
        async: false,
        success: async (resultData) => {

          var montaDiscussaoAntiga = "";

          console.log("Discussão", resultData);

          if (resultData.d.results.length > 0) {

            for (var i = 0; i < resultData.d.results.length; i++) {

              var thread = resultData.d.results[i].Thread;

              jQuery.ajax({
                url: `${this.props.siteurl}/_api/web/lists/getbytitle('DiscussaoAntiga')/items?$top=4999&$select=ID,Title,Area,Autor,Created,NotificarArea,Mensagem,Anexos&$filter=(substringof('${thread}', Thread)) and (Thread ne '${thread}')`,
                type: "GET",
                headers: { 'Accept': 'application/json; odata=verbose;' },
                async: false,
                success: async (resultData) => {

                  console.log("Mensagens", resultData);

                  if (resultData.d.results.length > 0) {

                    for (var x = 0; x < resultData.d.results.length; x++) {

                      var area = resultData.d.results[x].Area;

                      var area2 = "-";

                      if (area != null) {
                        area2 = area.replaceAll(/[^A-Z0-9]+/ig, "_");

                        if (_arrMontaBotoes.indexOf(area) == -1) _arrMontaBotoes.push(area);
                      }

                      var autor = resultData.d.results[x].Autor;
                      if (autor == "Gustavo Martins de Lima e Silva, Luis (ND)") autor = "Usuário da ";
                      else autor = resultData.d.results[x].Autor + " -";
                      var criado = new Date(resultData.d.results[x].Created);
                      var criado2 = new Date(criado.setHours(criado.getHours()));
                      var criadoHora = criado2.getHours() + ":" + ("0" + (criado2.getMinutes() + 1)).slice(-2) + ":" + criado2.getSeconds();
                      var criadoData = ("0" + criado2.getDate()).slice(-2) + '/' + ("0" + (criado2.getMonth() + 1)).slice(-2) + '/' + criado2.getFullYear();
                      var notificarArea = resultData.d.results[x].NotificarArea;
                      var mensagem = resultData.d.results[x].Mensagem;
                      var anexos = resultData.d.results[x].Anexos;
                      if (anexos == null) anexos = `<div class="text-secondary">Nenhum anexo encontrado.</div>`;

                      montaDiscussaoAntiga += `
                    <div class="p-0 mb-0 bg-light text-dark rounded comment ${area2}">
        
                    <div class="p-3 mb-2 alert-danger text-dark rounded-top ">
                    <b>Comentário postado por:</b> ${autor} Área ${area} em ${criadoData} às ${criadoHora}<br>
                    <b>Áreas notificadas:</b> ${notificarArea}
                    
                    </div>
                    <br/>
                    <div class="p-3">
                    ${mensagem}
        
                    Anexos relacionados:<br/>
                    ${anexos}
                    </div>
                    <br/><br/>
                    </div>
    
                `
                    }

                    $("#conteudoDiscusssaoAntiga").append(montaDiscussaoAntiga);
                  }



                },
                error: function (jqXHR, textStatus, errorThrown) {
                  console.log(jqXHR.responseText);
                }

              })

            }
          }


        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(jqXHR.responseText);
        }

      })

    }




    $("#conteudoDiscusssaoNova").empty();

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('ListaDiscussao')/items?$top=4999&$select=ID,Title,Author/Title,Area/Title,Created,NotificarArea/Title,Mensagem,txtAnexosRelacionados&$expand=Author,Area,NotificarArea&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        var montaDiscussaoNova = ``;
        var montaBotoes = ``;
        var idBtn = 0;

        if (resultData.d.results.length > 0) {

          //console.log("resultData.d.results.length", resultData.d.results.length);

          for (var i = 0; i < resultData.d.results.length; i++) {

            var montaLinksAnexos = ``;

            var autor = resultData.d.results[i].Author.Title;
            var area = resultData.d.results[i].Area.Title;
            var area2 = area.replaceAll(/[^A-Z0-9]+/ig, "_");
            var criado = new Date(resultData.d.results[i].Created);
            var criado2 = new Date(criado.setHours(criado.getHours()));
            var criadoHora = criado2.getHours() + ":" + ("0" + (criado2.getMinutes() + 1)).slice(-2) + ":" + criado2.getSeconds();
            var criadoData = ("0" + criado2.getDate()).slice(-2) + '/' + ("0" + (criado2.getMonth() + 1)).slice(-2) + '/' + criado2.getFullYear();
            var arrNotificarArea = [];
            var notificarArea = resultData.d.results[i].NotificarArea.results;
            var mensagem = resultData.d.results[i].Mensagem;

            for (var x = 0; x < notificarArea.length; x++) {
              arrNotificarArea.push(notificarArea[x].Title);
            }

            var anexosRelacionados = resultData.d.results[i].txtAnexosRelacionados;

            if (anexosRelacionados !== null) {

              var arrAnexosRelacionados = anexosRelacionados.split(',');

              for (var y = 0; y < arrAnexosRelacionados.length; y++) {

                montaLinksAnexos += `<a target="_blank" href="${_siteURL}/AnexosSAP/${_idProposta}/${arrAnexosRelacionados[y]}">${arrAnexosRelacionados[y]}</a><br/>`;

              }

            } else montaLinksAnexos = `<div class="text-secondary">Nenhum anexo encontrado.</div>`;

            idBtn++;

            if (_arrMontaBotoes.indexOf(area) == -1) _arrMontaBotoes.push(area);

            montaDiscussaoNova += `
            <div class="p-0 mb-0 bg-light text-dark rounded comment ${area2}">

            <div class="p-3 mb-2 alert-danger text-dark rounded-top ">
            <b>Comentário postado por:</b> ${autor} - Área ${area} em ${criadoData} às ${criadoHora}<br>
            <b>Áreas notificadas:</b> ${arrNotificarArea.toString()}
            
            </div>
            <br/>
            <div class="p-3">
            ${mensagem}

            Anexos relacionados:<br/>
            ${montaLinksAnexos}
            </div>
            <br/><br/>
            </div>

            `

          }

          var reactHandlerBotoesAreas = this;

          reactHandlerBotoesAreas.setState({
            itemsBotoes: _arrMontaBotoes
          });


        }

        //console.log("montaBotoes2", montaBotoes);
        $("#conteudoDiscusssaoNova").append(montaDiscussaoNova);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }

    })


    var reactHandlerBotoesAreas = this;

    reactHandlerBotoesAreas.setState({
      itemsBotoes: _arrMontaBotoes
    });


  }


  protected modalResponderDiscussao() {

    jQuery("#modalDiscussao").modal({ backdrop: 'static', keyboard: false });


  }

  private onTextChangeMensagemDiscussao = (newText: string) => {
    _mensagemDiscussao = newText;
    return newText;
  }

  private async cadastrarDiscussao() {

    $("#btnCadastrarDiscussao").prop("disabled", true);
    $("#btnCadastrarDiscussaoCancelar").prop("disabled", true);

    var area = jQuery('#ddlArea option:selected').val();
    var mensagem = _mensagemDiscussao;

    var arrAnexosSelecionados = [];
    $.each($("input[name='checkAnexosSelecionados']:checked"), function () {
      arrAnexosSelecionados.push($(this).val());
    });

    var arrNotificarArea = [];
    $.each($("input[name='checkNotificarArea']:checked"), function () {
      arrNotificarArea.push($(this).val());
    });

    var arrNomeNotificarArea = [];
    $.each($("input[name='checkNotificarArea']:checked"), function () {
      arrNomeNotificarArea.push($(this).attr('grupo'));
    });

    if (area == "0") {
      alert("Escolha a área!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      $("#btnCadastrarDiscussaoCancelar").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == "") {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      $("#btnCadastrarDiscussaoCancelar").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == "<p><br></p>") {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      $("#btnCadastrarDiscussaoCancelar").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == undefined) {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      $("#btnCadastrarDiscussaoCancelar").prop("disabled", false);
      return false;
    }


    if (arrNotificarArea.length == 0) {
      alert("Escolha uma área a ser notificada!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      $("#btnCadastrarDiscussaoCancelar").prop("disabled", false);
      return false;
    }

    await _web.lists
      .getByTitle("ListaDiscussao")
      .items.add({
        AreaId: area,
        Mensagem: mensagem,
        NotificarAreaId: { "results": arrNotificarArea },
        txtgrupoNotificaArea: arrNomeNotificarArea.toString(),
        PropostaId: _idProposta,
        SiteAntigo: "Nao",
        txtAnexosRelacionados: arrAnexosSelecionados.toString(),
      })
      .then(response => {

        //_idProposta = response.data.ID;
        console.log("gravou discussão");
        $("#modalDiscussao").modal('hide');
        $("#btnCadastrarDiscussao").prop("disabled", false);
        $("#ddlArea").val('0');
        $("input[name='checkAnexosSelecionados']").prop('checked', false);
        $("input[name='checkNotificarArea']").prop('checked', false)
        $(".ql-editor").empty();
        _mensagemDiscussao = "";

        this.getDiscussao();

      }).catch((error: any) => {
        console.log(error);
      });

  }

  private recarregaTarefas() {

    var reactHanthisdlerTarefas = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHanthisdlerTarefas.setState({
          itemsTarefas: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


  }

  protected modalReabrirProposta() {

    jQuery("#modalReabrirProposta").modal({ backdrop: 'static', keyboard: false });

  }

  protected async reabrirProposta() {

    $("#btnModalReabrirProposta").prop("disabled", true);
    $("#btnModalReabrirPropostaCancelar").prop("disabled", true);

    var novoStatus = $("#ddlStatusReabertura").val();
    var motivo = $("#ddlMotivoReabertura").val();
    var justificativa = $("#txtJustificativaReabertura").val();

    if (novoStatus == "0") {
      alert("Selecione o novo status da Proposta!");
      $("#btnModalReabrirProposta").prop("disabled", false);
      $("#btnModalReabrirPropostaCancelar").prop("disabled", false);
      return false;
    }

    if (motivo == "0") {
      alert("Selecione o motivo!");
      $("#btnModalReabrirProposta").prop("disabled", false);
      $("#btnModalReabrirPropostaCancelar").prop("disabled", false);
      return false;
    }

    if (justificativa == "") {
      alert("Forneça uma justificativa!");
      $("#btnModalReabrirProposta").prop("disabled", false);
      $("#btnModalReabrirPropostaCancelar").prop("disabled", false);
      return false;
    }

    console.log("novoStatus", novoStatus);
    console.log("motivo", motivo);
    console.log("justificativa", justificativa);

    if (novoStatus != "Voltar para em andamento") {
      var revisada = false;
    }
    else {
      var revisada = true;
      novoStatus = "Em análise";
    }

    await _web.lists
      .getByTitle("PropostasSAP")
      .items.getById(_idProposta).update({
        JustificativaFinal: justificativa,
        Status: novoStatus,
        Motivo: motivo,
        Revisada: revisada
      })
      .then(response => {

        if (novoStatus == "Em análise") {

          jquery.ajax({
            url: `${_siteURL}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$filter=Proposta/ID eq ` + _idProposta,
            type: "GET",
            async: false,
            headers: { 'Accept': 'application/json; odata=verbose;' },
            success: async function (resultData) {

              if (resultData.d.results.length > 0) {

                for (var i = 0; i < resultData.d.results.length; i++) {

                  var idTarefa = resultData.d.results[i].ID;

                  await _web.lists
                    .getByTitle("Tarefas")
                    .items.getById(idTarefa).update({
                      DataRealTermino: null,
                      Justificativa: "",
                      Status: "Em análise",
                    }).then(response => {
                      console.log("Atualizou a tarefa");
                    }).catch((error: any) => {
                      console.log(error);
                    });
                }
                $("#modalReabrirProposta").modal('hide');
                window.location.href = `Proposta-Detalhes.aspx?PropostasID=${_idProposta}`;
              }
            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(jqXHR.responseText);
            }
          });


        } else {
          $("#modalReabrirProposta").modal('hide');
          window.location.href = `Proposta-Detalhes.aspx?PropostasID=${_idProposta}`;
        }



      })


  }

  protected editarProposta() {

    window.location.href = `Propostas-SAP-Editar.aspx?PropostasID=${_idProposta}`;

  }

  protected voltar() {

    history.back();

  }

  protected upload() {


    var areaAnexo = $("#ddlAreaAnexo option:selected").text();

    if (areaAnexo == "Selecione...") {
      alert("Escolha a Área do anexo");
      return false;
    }

    var files = (document.querySelector("#input") as HTMLInputElement).files;

    if (files.length != 0) {

      jQuery("#modalCarregando").modal({ backdrop: 'static', keyboard: false });

      if (_pastaCriada != "Sim") {
        _web.lists.getByTitle("AnexosSAP").rootFolder.folders.add(`${_idProposta}`);
      }

      for (var i = 0; i < files.length; i++) {

        var nomeArquivo = files[i].name;
        var rplNomeArquivo = nomeArquivo.replace(/[^0123456789.,a-zA-Z]/g, '');

        _web.getFolderByServerRelativeUrl(`${_caminho}/AnexosSAP/${_idProposta}`)
          .files.add(rplNomeArquivo, files[i], true)
          .then(async data => {

            data.file.getItem().then(async item => {
              var idAnexo = item.ID;

              await _web.lists
                .getByTitle("AnexosSAP")
                .items.getById(idAnexo).update({
                  Area: areaAnexo,
                })
                .then(async response => {

                  if (i == files.length) {

                    console.log("anexou:" + rplNomeArquivo);
                    $("#conteudoLoading").modal('hide');
                    jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });

                  }

                }).catch(err => {
                  console.log("err", err);
                });

            });


          }).catch(err => {
            console.log("err", err);
          });

      }

      //const folderAddResult = _web.folders.add(`${_caminho}/Anexos/${_idProposta}`);

    } else {
      alert("Nenhum anexo encontrado!")
      return false;
    }

  }

  protected fecharSucesso() {

    $("#conteudoAnexo").empty();
    this.getAnexos();
    jQuery("#modalCarregando").modal('hide');
    jQuery("#modalSucesso").modal('hide');

  }

  protected filtrar(area) {

    console.log("filtrar area", area);

    if (area == "Todos") {

      $(".comment").css("display", "block");

    } else {

      $(".comment").hide();
      $('.' + area).css("display", "block");

    }


  }



}
