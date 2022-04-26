import * as React from 'react';
import styles from './SapNovaProposta.module.scss';
import { ISapNovaPropostaProps } from './ISapNovaPropostaProps';
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
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
//import "jquery-mask-plugin";

import InputMask from 'react-input-mask';
import { deprecationHandler } from 'moment';

//import 'jquery-mask-plugin/dist/jquery.mask.min';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;

export interface IReactGetItemsState {
  itemsRepresentante: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsClientes: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsProdutos: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsSegmento: [];
  itemsSetor: [];
  itemsModalidade: [];
  itemsInstalacao: [];
  itemsGarantia: [];
  itemsTipoGarantia: [];
  itemsPrazoGarantia: [];
  itemsOutrosServicos: [];

}

export default class SapNovaProposta extends React.Component<ISapNovaPropostaProps, IReactGetItemsState> {

  public constructor(props: ISapNovaPropostaProps, state: IReactGetItemsState) {
    super(props);
    this.state = {
      itemsRepresentante: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsClientes: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsProdutos: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsSegmento: [],
      itemsSetor: [],
      itemsModalidade: [],
      itemsInstalacao: [],
      itemsGarantia: [],
      itemsTipoGarantia: [],
      itemsPrazoGarantia: [],
      itemsOutrosServicos: [],
    };
  }


  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    document
      .getElementById("btnSalvar")
      .addEventListener("click", (e: Event) => this.salvar(false));

    document
      .getElementById("btnIniciarAprovacao")
      .addEventListener("click", (e: Event) => jQuery("#modalConfirmarIniciarFluxo").modal({ backdrop: 'static', keyboard: false }));


    document
      .getElementById("btIniciarFluxo")
      .addEventListener("click", (e: Event) => this.salvar(true));

    document
      .getElementById("btnSucesso")
      .addEventListener("click", (e: Event) => this.fecharSucesso());


    this.handler();



  }


  public render(): React.ReactElement<ISapNovaPropostaProps> {
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
                      <div className="form-group col-md-8">
                        <label htmlFor="txtTitulo">Tipo de análise</label><span className="required"> *</span>
                        <div className="form-check">
                          <input className="form-check-input" type="radio" name="checkTipoAnalise" id="checkTipoAnaliseProposta" />
                          <label className="form-check-label" htmlFor="checkTipoAnaliseProposta">
                            Proposta
                          </label>
                        </div>
                        <div className="form-check">
                          <input className="form-check-input" type="radio" name="checkTipoAnalise" id="checkTipoAnaliseContrato" />
                          <label className="form-check-label" htmlFor="checkTipoAnaliseContrato">
                            Contrato
                          </label>
                        </div>
                      </div>
                      <div className="form-group col-md-4">
                      </div>
                    </div>
                  </div>
                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-6">
                        <label htmlFor="txtSintese">Síntese</label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtSintese" />
                      </div>
                      <div className="form-group col-md-6">
                        <label htmlFor="txtIdentificacaoOportunidade">Identificação da Oportunidade </label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtIdentificacaoOportunidade" />
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataEntregaPropostaCliente">Data da entrega da Proposta ao Cliente</label><span className="required"> *</span>
                        <InputMask mask="99/99/9999" className="form-control" maskChar={null} id="dtDataEntregaPropostaCliente" />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataFinalQuestionamentos">Data final de questionamentos</label><span className="required"> *</span>
                        <InputMask mask="99/99/9999" className="form-control" maskChar={null} id="dtDataFinalQuestionamentos" />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataValidadeProposta">Data de validade da Proposta</label><span className="required"> *</span>
                        <InputMask mask="99/99/9999" className="form-control" maskChar={null} id="dtDataValidadeProposta" />
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlRepresentante">Representante</label><span className="required"> *</span>
                        <select id="ddlRepresentante" className="form-control">
                          <option value="0" selected>Selecione...</option>
                          {this.state.itemsRepresentante.map(function (item, key) {
                            return (
                              <option value={item.ID}>{item.Title}</option>
                            );
                          })}
                        </select>
                      </div>
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlCliente">Cliente </label><span className="required"> *</span>
                        <select id="ddlCliente" className="form-control">
                          <option value="0" selected>Selecione...</option>
                          {this.state.itemsClientes.map(function (item, key) {
                            return (
                              <option value={item.ID}>{item.Title}</option>
                            );
                          })}
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="txtPropostaRevisadaReferencia">Proposta revisada/referência</label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtPropostaRevisadaReferencia" />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtSST">SST</label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtSST" />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtCondicoesPagamento">Condições de pagamento </label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtCondicoesPagamento" />
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
                    <textarea id="txtDadosProposta" className="form-control" rows={4}></textarea>
                  </div>

                  <div className="form-group">
                    <label htmlFor="txtJustificativa">Justificativa</label> <span className="required"> *</span>
                    <textarea id="txtJustificativa" className="form-control" rows={4}></textarea>
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
                        <label htmlFor="txtTitulo">Segmento</label><span className="required"> *</span>
                        {this.state.itemsSegmento.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkSegmento' type="checkbox" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtTitulo">Setor</label><span className="required"> *</span>

                        {this.state.itemsSetor.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" type="radio" name="checkSetor" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="ddlModalidade">Modalidade </label><span className="required"> *</span>

                        {this.state.itemsModalidade.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkModalidade' type="radio" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <label htmlFor="txtNumeroEditalRFPRFQRFI">Número do Edital, RFP, RFQ ou RFI </label> <span className="required"> *</span>
                    <input type="text" className="form-control" id="txtNumeroEditalRFPRFQRFI" />
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
                        <label htmlFor="txtQuantidade">Quantidade</label><span className="required"> *</span>
                        <input type="number" style={{ "width": "120px" }} className="form-control" id="txtQuantidade" />
                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtTitulo">Instalação</label><span className="required"> *</span><br></br>
                        {this.state.itemsInstalacao.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkInstalacao' type="radio" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}
                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtTitulo">Garantia</label><span className="required"> *</span><br></br>

                        {this.state.itemsGarantia.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkGarantia' type="radio" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}


                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtTitulo">Tipo de garantia </label><span className="required"> *</span><br></br>

                        {this.state.itemsTipoGarantia.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkTipoGarantia' type="radio" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-2">
                        <label htmlFor="txtTitulo">Prazo de garantia </label><span className="required"> *</span>

                        {this.state.itemsPrazoGarantia.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkPrazoGarantia' type="radio" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                      <div className="form-group col-md-2">
                        <label htmlFor="txtTitulo">Outros serviços</label><span className="required"> *</span><br></br>

                        {this.state.itemsOutrosServicos.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkOutrosServicos' type="checkbox" value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                      <div className="form-group col-md-8">
                        <label htmlFor="ddlProduto">Produto</label><span className={styles.required}> *</span>
                        <table>
                          <tr>
                            <td>
                              <div className="col-sm-2">
                                <select multiple={true} id='ddlProduto1' className="form-control" name="ddlProduto1" style={{ "height": "194px", "width": "200px" }}>

                                  {this.state.itemsProdutos.map(function (item, key) {
                                    return (
                                      <option className="optProduto" value={item.ID}>{item.Title}</option>
                                    );
                                  })}

                                </select>
                              </div>
                            </td>
                            <td>
                              <div>
                                <input type="button" className="btn btn-light" id="addButtonProduto" onClick={this.addButtonProduto} value="Adicionar >" /></div><br />
                              <input type="button" className="btn btn-light" id="removeButtonProduto" onClick={this.removeButtonProduto} value="< Remover" />
                            </td>
                            <td>
                              <div className="col-sm-2">
                                <select multiple={true} id="ddlProduto2" className="form-control" name="ddlProduto2" style={{ "height": "194px", "width": "200px" }}>
                                </select>
                              </div>
                            </td>
                          </tr>
                        </table>
                      </div>
                    </div>
                  </div>


                </div>
              </div>
            </div>




            <div className="card">
              <div className="card-header btn" id="headingProduto" data-toggle="collapse" data-target="#collapseProduto" aria-expanded="true" aria-controls="collapseProduto">
                <h5 className="mb-0 text-info" >
                  Áreas Responsáveis pela Proposta
                </h5>
              </div>
              <div id="collapseProduto" className="collapse show" aria-labelledby="headingOne" >

                <div className="card-body">

                  <label htmlFor="ddlProduto">Áreas</label><span className={styles.required}> *</span>
                  <table>
                    <tr>
                      <td>
                        <div className="col-sm-6">
                          <select multiple={true} id='ddlProduto1' className="form-control" name="ddlProduto1" style={{ "height": "194px", "width": "350px" }}>
                            <option className="optProduto" value="Adequação civil">Adequação civil</option>
                          </select>
                        </div>
                      </td>
                      <td>
                        <div>
                          <input type="button" className="btn btn-light" id="addButtonProduto" value="Adicionar >" alt="Salvar" /></div><br />
                        <input type="button" className="btn btn-light" id="removeButtonProduto" value="< Remover"
                          alt="Salvar" />
                      </td>
                      <td>
                        <div className="col-sm-6">
                          <select multiple={true} id="ddlProduto2" className="form-control" name="ddlProduto2" style={{ "height": "194px", "width": "350px" }}>
                          </select>
                        </div>
                      </td>
                    </tr>
                  </table>
                </div>
              </div>
            </div>

            <br></br>

            <div className="text-right">
              <button id="btnSalvar" className="btn btn-secondary">Salvar</button>&nbsp;
              <button id="btnIniciarAprovacao" className="btn btn-success" >Enviar para Aprovação</button>
            </div>


          </div>
        </div>



        <div className="modal fade" id="modalConfirmarIniciarFluxo" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">
                Deseja realmente iniciar o fluxo?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btIniciarFluxo" type="button" className="btn btn-primary">Iniciar fluxo</button>
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
                Proposta cadastrada com sucesso!
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

  protected handler() {

    var reactHandlerRepresentante = this;
    var reactHandlerClientes = this;
    var reactHandlerSegmento = this;
    var reactHandlerSetor = this;
    var reactHandlerModalidade = this;
    var reactHandlerInstalacao = this;
    var reactHandlerGarantia = this;
    var reactHandlerTipoGarantia = this;
    var reactHandlerPrazoGarantia = this;
    var reactHandlerOutrosServicos = this;
    var reactHandlerProdutos = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Representantes')/items?$top=4999&$filter=Ativo eq 1&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerRepresentante.setState({
          itemsRepresentante: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Clientes')/items?$top=4999&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerClientes.setState({
          itemsClientes: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'Segmento'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerSegmento.setState({
          itemsSegmento: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'Setor'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerSetor.setState({
          itemsSetor: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'Modalidade'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerModalidade.setState({
          itemsModalidade: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'Instalacao'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerInstalacao.setState({
          itemsInstalacao: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'Garantia'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerGarantia.setState({
          itemsGarantia: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'TipoGarantia'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerTipoGarantia.setState({
          itemsTipoGarantia: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'TipoGarantia'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerTipoGarantia.setState({
          itemsTipoGarantia: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'PrazoGarantia'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerPrazoGarantia.setState({
          itemsPrazoGarantia: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('PropostasSAP')/fields?$filter=EntityPropertyName eq 'OutrosServicos'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerOutrosServicos.setState({
          itemsOutrosServicos: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Produtos')/items?$top=4999&$filter=Ativo eq 1&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerProdutos.setState({
          itemsProdutos: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });




  }

  protected addButtonProduto = () => {
    var $options = $('#ddlProduto1 option:selected');
    $options.appendTo("#ddlProduto2");
  }

  protected removeButtonProduto = () => {
    var $options = $('#ddlProduto2 option:selected');
    $options.appendTo("#ddlProduto1");
  }


  protected salvar(iniciarFluxo) {

    $("#modalConfirmarIniciarFluxo").modal('hide');

    console.log("entro no salvar")

    var tipoAnaliseProposta;

    if ($('#checkTipoAnaliseProposta').is(':checked')) { tipoAnaliseProposta = "Proposta" };
    if ($('#checkTipoAnaliseContrato').is(':checked')) { tipoAnaliseProposta = "Contrato" };

    var sintese = $("#txtSintese").val();
    var identificacaoOportunidade = $("#txtIdentificacaoOportunidade").val();

    var dataEntregaPropostaCliente = "" + jQuery("#dtDataEntregaPropostaCliente").val() + "";
    var dataEntregaPropostaClienteDia = dataEntregaPropostaCliente.substring(0, 2);
    var dataEntregaPropostaClienteMes = dataEntregaPropostaCliente.substring(3, 5);
    var dataEntregaPropostaClienteAno = dataEntregaPropostaCliente.substring(6, 10);
    var formDataEntregaPropostaCliente = dataEntregaPropostaClienteAno + "-" + dataEntregaPropostaClienteMes + "-" + dataEntregaPropostaClienteDia;

    var dataFinalQuestionamentos = "" + jQuery("#dtDataFinalQuestionamentos").val() + "";
    var dataFinalQuestionamentosDia = dataFinalQuestionamentos.substring(0, 2);
    var dataFinalQuestionamentosMes = dataFinalQuestionamentos.substring(3, 5);
    var dataFinalQuestionamentosAno = dataFinalQuestionamentos.substring(6, 10);
    var formDataFinalQuestionamentos = dataFinalQuestionamentosAno + "-" + dataFinalQuestionamentosMes + "-" + dataFinalQuestionamentosDia;

    var dataValidadeProposta = "" + jQuery("#dtDataValidadeProposta").val() + "";
    var dataValidadePropostaDia = dataValidadeProposta.substring(0, 2);
    var dataValidadePropostaMes = dataValidadeProposta.substring(3, 5);
    var dataValidadePropostaAno = dataValidadeProposta.substring(6, 10);
    var formDataValidadeProposta = dataValidadePropostaAno + "-" + dataValidadePropostaMes + "-" + dataValidadePropostaDia;

    console.log("sintese", sintese);
    console.log("tipoAnaliseProposta", tipoAnaliseProposta);
    console.log("identificacaoOportunidade", identificacaoOportunidade);
    console.log("formDataEntregaPropostaCliente", formDataEntregaPropostaCliente);
    console.log("formDataFinalQuestionamentos", formDataFinalQuestionamentos);
    console.log("formDataValidadeProposta", formDataValidadeProposta);

    var representante = $("#ddlRepresentante").val();
    var cliente = $("#ddlCliente").val();
    var propostaRevisadaReferencia = $("#txtPropostaRevisadaReferencia").val();
    var SST = $("#txtSST").val();
    var condicoesPagamento = $("#txtCondicoesPagamento").val();
    var dadosProposta = $("#txtDadosProposta").val();
    var justificativaFinal = $("#txtJustificativa").val();

    var arrSegmento = [];
    $.each($("input[name='checkSegmento']:checked"), function () {
      arrSegmento.push($(this).val());
    });

    var arrSetor = [];
    $.each($("input[name='checkSetor']:checked"), function () {
      arrSetor.push($(this).val());
    });

    var arrModalidade = [];
    $.each($("input[name='checkModalidade']:checked"), function () {
      arrModalidade.push($(this).val());
    });

    var numeroEditalRFPRFQRFI = $("#txtNumeroEditalRFPRFQRFI").val();
    var quantidade = $("#txtQuantidade").val();

    var arrInstalacao  = [];
    $.each($("input[name='checkInstalacao']:checked"), function () {
      arrInstalacao.push($(this).val());
    });

    
    _web.lists
      .getByTitle("PropostasSAP")
      .items.add({
        Title: sintese,
        TipoAnalise: tipoAnaliseProposta,
        IdentificacaoOportunidade: identificacaoOportunidade,
        DataEntregaPropostaCliente: formDataEntregaPropostaCliente,
        DataFinalQuestionamentos: formDataFinalQuestionamentos,
        DataValidadeProposta: formDataValidadeProposta,
        RepresentanteId: representante,
        ClienteId: cliente,
        PropostaRevisadaReferencia: propostaRevisadaReferencia,
        SST: SST,
        CondicoesPagamento: condicoesPagamento,
        DadosProposta: dadosProposta,
        JustificativaFinal: justificativaFinal,
        Segmento: { "results": arrSegmento },
        Setor: arrSegmento[0],
        Modalidade: arrModalidade[0],
        NumeroEditalRFPRFQRFI: numeroEditalRFPRFQRFI,
        Quantidade: quantidade,
        Instalacao: arrInstalacao[0]
      })
      .then(response => {
        console.log("Gravou!!");
        jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })
      }).catch((error: any) => {
        console.log(error);

      });

  }


  fecharSucesso() {

    $("#modalSucesso").modal('hide');
    window.location.href = `Nova-Proposta-SAP.aspx`;

  }




}
