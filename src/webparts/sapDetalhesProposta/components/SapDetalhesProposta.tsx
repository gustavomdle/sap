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
import { first } from 'lodash';

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
    };
  }


  public async componentDidMount() {

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idProposta = parseInt(queryParms.getValue("PropostasID"));

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    _siteURL = this.props.siteurl;

    //let groups = await _web.currentUser.groups();

    await _web.currentUser.get().then(f => {
      console.log("user", f);
      var id = f.Id;

      var grupos = [];

      jQuery.ajax({
        url: `${this.props.siteurl}/_api/web/GetUserById(${id})/Groups`,
        type: "GET",
        headers: { 'Accept': 'application/json; odata=verbose;' },
        async: false,
        success: async function (resultData) {

          console.log("resultDataGrupo", resultData);

          if (resultData.d.results.length > 0) {

            for (var i = 0; i < resultData.d.results.length; i++) {

              grupos.push(resultData.d.results[i].Title);

            }

          }

        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(textStatus);
        }

      })

      console.log("grupos", grupos);
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



    var reactHandlerRepresentante = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerRepresentante.setState({
          itemsTarefas: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(textStatus);
      }
    });


    this.getProposta();
    this.getTarefas();
    this.getAnexos();
    this.getSelecaoAreas();
    this.getDiscussaoNova();

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
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="txtNumeroProposta">Número da Proposta</label><br></br>
                        <span className="text-info" id='txtNumeroProposta'></span>
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtStatus">Status</label><br></br>
                        <span className="text-info" id='txtStatus'></span>
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtTipoAnalise">Tipo de análise</label><br></br>
                        <span className="text-info" id='txtTipoAnalise'></span>
                      </div>
                    </div>
                  </div>




                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-7">
                        <label htmlFor="txtSintese">Síntese</label><br></br>
                        <span className="text-info" id='txtSintese'></span>
                      </div>
                      <div className="form-group col-md-5">
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

            <div className="card">
              <div className="card-header btn" id="headingDiscussao" data-toggle="collapse" data-target="#collapseDiscussao" aria-expanded="true" aria-controls="collapseDiscussao">
                <h5 className="mb-0 text-info" >
                  Discussão
                </h5>
              </div>
              <div id="collapseDiscussao" className="collapse show" aria-labelledby="headingOne" >

                <div className="card-body">

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

                        console.log("item.Status", item.Status);

                        if (item.Status == "Em análise") {

                          console.log("item.Title", item.Title);
                          console.log("_grupos.indexOf(item.Title)", _grupos.indexOf(item.Title));

                          if (_grupos.indexOf(item.Title) !== -1) {

                            return (

                              <><tr>
                                <td>{item.Title}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td>{vlrJustificativa}</td>
                                <td><button onClick={() => this.detalhesTarefas(item.ID, item.GrupoSharepoint.Title)} type="button" className="btn btn-success btn-sm">Detalhes</button></td>
                              </tr></>
                            );

                          } else {

                            return (

                              <><tr>
                                <td>{item.Title}</td>
                                <td>{item.Status}</td>
                                <td>{dtDataPlanejadaTermino}</td>
                                <td>{dtDataRealTermino}</td>
                                <td>{vlrJustificativa}</td>
                                <td></td>
                              </tr></>
                            );

                          }


                        } else {

                          return (

                            <><tr>
                              <td>{item.Title}</td>
                              <td>{item.Status}</td>
                              <td>{dtDataPlanejadaTermino}</td>
                              <td>{dtDataRealTermino}</td>
                              <td>{vlrJustificativa}</td>
                              <td></td>
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

            <div className="text-right">
              <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
            </div>


          </div>

        </div>

        <div className="modal fade" id="modalDiscussao" tabIndex={-1} role="dialog" aria-labelledby="discussaoTitle" aria-hidden="true">
          <div className="modal-dialog modal-dialog-scrollable" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="discussaoTitle">Cadastrar Discussão</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
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
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
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
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
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
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button type="button" id='btnAprovarTarefa' className="btn btn-primary">Aprovar</button>
              </div>
            </div>
          </div>
        </div>

      </>

    );
  }



  protected getProposta() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('PropostasSAP')/items?$select=ID,Title,TipoAnalise,IdentificacaoOportunidade,DataEntregaPropostaCliente,DataFinalQuestionamentos,DataValidadeProposta,Representante/ID,Representante/Title,Cliente/ID,Cliente/Title,PropostaRevisadaReferencia,CondicoesPagamento,DadosProposta,Segmento,Setor,Modalidade,NumeroEditalRFPRFQRFI,Instalacao,Quantidade,Garantia,TipoGarantia,PrazoGarantia,OutrosServicos,Produto/ID,Produto/Title,Numero,Status&$expand=Representante,Cliente,Produto&$filter=ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        //console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

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



          }

        }



      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(textStatus);
      }



    })

  }


  protected getTarefas() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        var arrAreas = [];
        var montaNotificarArea = "";
        var montaTarefas = "";

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            arrAreas.push(resultData.d.results[i].GrupoSharepoint.Title);
            var id = resultData.d.results[i].ID;
            var dataPlanejadaTermino = new Date(resultData.d.results[i].DataPlanejadaTermino);
            var dtDataPlanejadaTermino = ("0" + dataPlanejadaTermino.getDate()).slice(-2) + '/' + ("0" + (dataPlanejadaTermino.getMonth() + 1)).slice(-2) + '/' + dataPlanejadaTermino.getFullYear();
            var dataRealTermino = new Date(resultData.d.results[i].DataRealTermino);
            var dtDataRealTermino = ("0" + dataRealTermino.getDate()).slice(-2) + '/' + ("0" + (dataRealTermino.getMonth() + 1)).slice(-2) + '/' + dataRealTermino.getFullYear();
            if (dtDataRealTermino == "31/12/1969") dtDataRealTermino = "-";
            var justificativa = resultData.d.results[i].Justificativa;

            console.log("justificativa", justificativa);

            if (justificativa == null) justificativa = "-";

          }

        }

        jQuery("#txtAreas").html(arrAreas.toString());

        for (i = 0; i < arrAreas.length; i++) {

          montaNotificarArea += `<div class="form-check">
                                    <input class="form-check-input" name="checkNotificarArea" grupo="${resultData.d.results[i].GrupoSharepoint.Title}" type="checkbox" value="${resultData.d.results[i].GrupoSharepoint.ID}">
                                    <label class="form-check-label">${resultData.d.results[i].GrupoSharepoint.Title}</label>
                                    </div>`;



        }

        $("#conteudoNotificarArea").append(montaNotificarArea);



      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(textStatus);
      }



    })


  }

  protected detalhesTarefas(id, nomeGrupo) {

    console.log("entrou no detalhesTarefas" + id);
    _idTarefa = id;

    jQuery("#txtModalArea").html(nomeGrupo);
    jQuery("#modalAprovacao").modal({ backdrop: 'static', keyboard: false });

  }

  protected async aprovar() {

    var status = jQuery('#ddlStatus option:selected').val();
    var justificativa = jQuery('#txtJustificativa').val();

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
      .then(response => {


        if (status == "Aprovada") {

          jQuery.ajax({
            url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title&$filter=(Proposta/ID eq ${_idProposta}) and (Status ne 'Aprovada' or Status ne 'Não envolve a Área')`,
            type: "GET",
            headers: { 'Accept': 'application/json; odata=verbose;' },
            async: false,
            success: async (resultData) => {

              if (resultData.d.results.length > 0) {

                console.log("Aprovou!!!");
                this.recarregaTarefas();
                $("#modalAprovacao").modal('hide');

              } else {

                await _web.lists
                  .getByTitle("PropostasSAP")
                  .items.getById(_idProposta).update({
                    Status: "Aprovado",
                  })
                  .then(response => {

                    console.log("Aprovou!!!");
                    this.recarregaTarefas();
                    $("#modalAprovacao").modal('hide');

                  })

              }


            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(textStatus);
            }

          })
          //se sim atualiza tarefas e proposta
        }


        else if (status == "Não envolve a Área") {

          jQuery.ajax({
            url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title&$filter=(Proposta/ID eq ${_idProposta}) and (Status ne 'Aprovada' or Status ne 'Não envolve a Área')`,
            type: "GET",
            headers: { 'Accept': 'application/json; odata=verbose;' },
            async: false,
            success: async (resultData) => {

              if (resultData.d.results.length > 0) {

                console.log("Aprovou!!!");
                this.recarregaTarefas();
                $("#modalAprovacao").modal('hide');

              } else {

                await _web.lists
                  .getByTitle("PropostasSAP")
                  .items.getById(_idProposta).update({
                    Status: "Aprovado",
                  })
                  .then(response => {

                    console.log("Aprovou!!!");
                    this.recarregaTarefas();
                    $("#modalAprovacao").modal('hide');

                  })

              }

            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(textStatus);
            }

          })


        }

        else if (status == "Reprovada") {

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

                    })

                }

                console.log("Aprovou!!!");
                this.recarregaTarefas();
                $("#modalAprovacao").modal('hide');

              }

            },
            error: function (jqXHR, textStatus, errorThrown) {
              console.log(textStatus);
            }

          })
        }


      }).catch((error: any) => {
        console.log(error);
      });

  }

  protected getAnexos() {

    //get anexos da biblioteca

    var montaAnexo = "";
    var montaAnexosRelacionados = "";

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Proposta-Detalhes.aspx", "");

    var idItem = 0;

    _web.getFolderByServerRelativeUrl(`${strRelativeURL}/AnexosSAP/${_idProposta}`)
      .expand("Folders, Files, ListItemAllFields").get().then(r => {
        //console.log("r", r);
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
          $("#conteudoAnexoNaoEncontrado").hide();
          montaAnexo = `<a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a>&nbsp; <br/>`

          $("#conteudoAnexo").append(montaAnexo);

          montaAnexosRelacionados = `<div class="form-check">
                                    <input class="form-check-input" name="checkAnexosSelecionados" type="checkbox" value="${item.Name}">
                                    <label class="form-check-label">${item.Name}</label>
                                    </div>`

          $("#conteudoAnexosRelacionados").append(montaAnexosRelacionados);


        })

      }).catch((error: any) => {
        console.log("Erro Anexo da biblioteca: ", error);
      });

    //fim anexos da biblioteca

  }

  protected getSelecaoAreas() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('AreasDiscussao')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title&$expand=GrupoSharepoint&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        var montaArea = `<option value="0" selected>Selecione...</option>`;

        // console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            montaArea += `<option value=${resultData.d.results[i].GrupoSharepoint.ID}>${resultData.d.results[i].GrupoSharepoint.Title}</option>`;

          }

        }

        $("#ddlArea").append(montaArea);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(textStatus);
      }

    })


  }

  protected getDiscussaoNova() {

    $("#conteudoDiscusssaoNova").empty();

    console.log("this.props.siteurl", this.props.siteurl);

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('ListaDiscussao_1911')/items?$select=ID,Title,Author/Title,Area/Title,Created,NotificarArea/Title,Mensagem,txtAnexosRelacionados&$expand=Author,Area,NotificarArea`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async function (resultData) {

        var montaDiscussaoNova = ``;

        console.log("resultData Discussao", resultData);

        if (resultData.d.results.length > 0) {

          console.log("resultData.d.results.length", resultData.d.results.length);

          for (var i = 0; i < resultData.d.results.length; i++) {

            var montaLinksAnexos = ``;

            console.log("entrou discussao");

            var autor = resultData.d.results[i].Author.Title;
            var area = resultData.d.results[i].Area.Title;
            var criado = new Date(resultData.d.results[i].Created);
            var criado2 = new Date(criado.setHours(criado.getHours() + 3));
            var criadoHora = criado2.getHours() + ":" + criado2.getMinutes() + ":" + criado2.getSeconds();
            var criadoData = ("0" + criado2.getDate()).slice(-2) + '/' + ("0" + (criado2.getMonth() + 1)).slice(-2) + '/' + criado2.getFullYear();
            var arrNotificarArea = [];
            var notificarArea = resultData.d.results[i].NotificarArea.results;
            var mensagem = resultData.d.results[i].Mensagem;

            for (var x = 0; x < notificarArea.length; x++) {
              arrNotificarArea.push(notificarArea[x].Title)
            }

            //var mensagem = resultData.d.results[i].Title;

            //console.log("Autor", autor);
            //console.log("Area", area);
            //console.log("criadoHora", criadoHora);
            //console.log("criadoData", criadoData);
            //console.log("notificarArea", notificarArea);
            //console.log("arrNotificarArea", arrNotificarArea.toString());
            // console.log("mensagem", mensagem);

            var anexosRelacionados = resultData.d.results[i].txtAnexosRelacionados;

            if (anexosRelacionados !== null) {

              var arrAnexosRelacionados = anexosRelacionados.split(',');

              for (var y = 0; y < arrAnexosRelacionados.length; y++) {

                montaLinksAnexos += `<a target="_blank" href="${_siteURL}/AnexosSAP/${_idProposta}/${arrAnexosRelacionados[y]}">${arrAnexosRelacionados[y]}</a><br/>`;

              }

            } else montaLinksAnexos = `<div class="text-secondary">Nenhum anexo encontrado.</div>`;


            console.log("arrAnexosRelacionados", arrAnexosRelacionados);

            montaDiscussaoNova += `
            <div class="p-0 mb-0 bg-light text-dark rounded ">

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
            <br/>
            </div>
            <br/>
            `


          }

        }

        $("#conteudoDiscusssaoNova").append(montaDiscussaoNova);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log("Erro em getDiscussaoNova: " + textStatus);
      }



    })




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

    console.log("_mensagemDiscussao", _mensagemDiscussao);


    if (area == "0") {
      alert("Escolha a área!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == "") {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == "<p><br></p>") {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      return false;
    }

    if (_mensagemDiscussao == undefined) {
      alert("Forneça uma mensagem!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      return false;
    }




    if (arrNotificarArea.length == 0) {
      alert("Escolha uma área a ser notificada!");
      $("#btnCadastrarDiscussao").prop("disabled", false);
      return false;
    }

    //console.log("area", area);
    //console.log("mensagem", mensagem);
    //console.log("arrAnexosSelecionados", arrAnexosSelecionados);
    //console.log("arrNotificarArea", arrNotificarArea);
    //console.log("arrNomeNotificarArea", arrNomeNotificarArea);



    await _web.lists
      .getByTitle("ListaDiscussao_1911")
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

        this.getDiscussaoNova();

      }).catch((error: any) => {
        console.log(error);
      });

  }

  private recarregaTarefas() {

    var reactHanthisdlerRepresentante = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHanthisdlerRepresentante.setState({
          itemsTarefas: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(textStatus);
      }
    });


  }


}
