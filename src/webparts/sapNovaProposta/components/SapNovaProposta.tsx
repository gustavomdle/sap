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
import { ICamlQuery } from '@pnp/sp/lists';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement, DatePicker } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';



//import * as Moment from 'moment';
//import 'react-dates/initialize';
//import 'react-dates/lib/css/_datepicker.css';
//import { DateRangePicker, SingleDatePicker, DayPickerRangeController } from 'react-dates';
//import { SPComponentLoader } from '@microsoft/sp-loader';
//import 'jqueryui'
//import 'jqueryui';
//import 'jquery-ui-bundle';


//import "jquery-mask-plugin";

import InputMask from 'react-input-mask';
import { deprecationHandler } from 'moment';

const divStyle = {
  padding: '0 0 0 20px'
};

//import 'jquery-mask-plugin/dist/jquery.mask.min';
require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;
var _criou = false;
var _arrAreaId = [];
var _arrAreaTexto = [];
var _caminho;
var _idProposta;
var _size: number = 0;

export interface ISamplePnPDateTimePickerWebpartState {

}

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
  itemsAreas: [
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
  startDate: any;
  endDate: any;
  focusedInput: any;

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
      itemsAreas: [
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
      startDate: "",
      endDate: "",
      focusedInput: "any",
    };
  }


  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    console.log("_caminho", _caminho);

    // jQuery("#dtDataEntrPropostaCliente").datepicker();

    document
      .getElementById("btnIniciarAprovacao")
      .addEventListener("click", (e: Event) => this.validar());


    document
      .getElementById("btIniciarFluxo")
      .addEventListener("click", (e: Event) => this.salvar());

    document
      .getElementById("btnSucesso")
      .addEventListener("click", (e: Event) => this.fecharSucesso());


    //var $options = $('#ddlProduto1 option:selected');
    //$options.appendTo("#ddlProduto2");


    $("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/Images1/loading.gif'/>
    <br/>Aguarde....<br/><br/>
    Dependendo do tamanho do anexo e a velocidade<br>
     da Internet essa ação pode demorar um pouco. <br>
     Não fechar a janela!<br/><br/>`);


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
                      <div className="form-group col-md-9">
                        <label htmlFor="txtSintese">Síntese</label><span className="required"> *</span>
                        <input type="text" className="form-control" id="txtSintese" />
                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtIdentificacaoOportunidade">Identificação da Oportunidade </label><span className="required"> *</span>
                        <InputMask mask="F999999" className="form-control" maskChar={null} id="txtIdentificacaoOportunidade" />
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataEntregaPropostaCliente">Data da entrega da Proposta ao Cliente</label><span className="required"> *</span>
                        <DatePicker minDate={this.addDaysWRONG()} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="form-control" id='dtDataEntregaPropCliente' />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataFinalQuestionamentos">Data final de questionamentos</label>
                        <DatePicker minDate={new Date()} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="form-control" id='dtDataFinalQuestionamentos' />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataValidadeProposta">Data de validade da Proposta</label>
                        <DatePicker minDate={new Date()} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="form-control" id='dtDataValidadeProposta' />
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
                      <div className="form-group col-md-8">
                        <label htmlFor="txtPropostaRevisadaReferencia">Proposta revisada/referência</label>
                        <input type="text" className="form-control" id="txtPropostaRevisadaReferencia" />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="txtCondicoesPagamento">Condições de pagamento </label>
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
                    <label htmlFor="txtNumeroEditalRFPRFQRFI">Número do Edital, RFP, RFQ ou RFI </label>
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
                        <label htmlFor="txtTitulo">Tipo de garantia </label><br></br>

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
                        <label htmlFor="txtTitulo">Outros serviços</label><br></br>

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
              <div className="card-header btn" id="headingArea" data-toggle="collapse" data-target="#collapseArea" aria-expanded="true" aria-controls="collapseArea">
                <h5 className="mb-0 text-info" >
                  Áreas Responsáveis pela Proposta
                </h5>
              </div>
              <div id="collapseArea" className="collapse show" aria-labelledby="headingOne" >

                <div className="card-body">

                  <label htmlFor="ddlProduto">Áreas</label><span className="required"> *</span>
                  <table>
                    <tr>
                      <td>
                        <div className="col-sm-6">
                          <select multiple={true} id='ddlArea1' className="form-control" name="ddlArea1" style={{ "height": "194px", "width": "350px" }}>

                            {this.state.itemsAreas.map(function (item, key) {
                              return (
                                <option className="optArea" value={item.ID}>{item.Title}</option>
                              );

                            })}

                          </select>
                        </div>
                      </td>
                      <td>
                        <div>
                          <input type="button" onClick={this.addButtonArea} className="btn btn-light" id="addButtonArea" value="Adicionar >" alt="Salvar" /></div><br />
                        <input type="button" onClick={this.removeButtonArea} className="btn btn-light" id="removeButtonArea" value="< Remover"
                          alt="Salvar" />
                      </td>
                      <td>
                        <div className="col-sm-6">
                          <select multiple={true} id="ddlArea2" className="form-control" name="ddlArea2" style={{ "height": "194px", "width": "350px" }}>
                          </select>
                        </div>
                      </td>
                    </tr>
                  </table>
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

                  <label htmlFor="ddlProduto">Anexo</label><span className="required"> *</span><br />
                  <input className="multi" data-maxsize="1024" type="file" id="input" multiple />

                </div>
              </div>
            </div>

            <br></br>

            <div className="text-right">
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

        <div className="modal fade" id="modalCarregando" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div>
            <div className="modal-dialog" role="document">
              <div className="modal-content">
                <div id='conteudoLoading' className='carregando'></div>
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

  private onFormatDate = (date: Date): string => {
    //return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
    return ("0" + date.getDate()).slice(-2) + '/' + ("0" + (date.getMonth() + 1)).slice(-2) + '/' + date.getFullYear();
  };


  private addDaysWRONG() {

    var date = new Date();
    var result = new Date();
    result.setDate(date.getDate() + 5);
    return result;
  }


  protected async handler() {

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
    var reactHandlerAreas = this;

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
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerProdutos.setState({
          itemsProdutos: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Areas')/items?$top=4999&$filter=Ativo eq 1&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerAreas.setState({
          itemsAreas: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    setTimeout(function () {

      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Comercial"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Jurídico"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Representante"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Propostas"; }).prop('selected', true);
      var $options = $('#ddlArea1 option:selected');
      $options.appendTo("#ddlArea2");


    }, 2000);

  }

  protected addButtonProduto = () => {
    var $options = $('#ddlProduto1 option:selected');
    $options.appendTo("#ddlProduto2");
  }

  protected removeButtonProduto = () => {
    var $options = $('#ddlProduto2 option:selected');
    $options.appendTo("#ddlProduto1");
  }

  protected addButtonArea = () => {
    var $options = $('#ddlArea1 option:selected');
    $options.appendTo("#ddlArea2");
  }

  protected removeButtonArea = () => {
    var $options = $('#ddlArea2 option:selected');
    $options.appendTo("#ddlArea1");
  }

  protected validar() {

    console.log("Entrou na validação");

    var tipoAnaliseProposta = "";

    if ($('#checkTipoAnaliseProposta').is(':checked')) { tipoAnaliseProposta = "Proposta" };
    if ($('#checkTipoAnaliseContrato').is(':checked')) { tipoAnaliseProposta = "Contrato" };

    var sintese = $("#txtSintese").val();
    var identificacaoOportunidade = $("#txtIdentificacaoOportunidade").val();
    var dataEntregaPropostaCliente = "" + jQuery("#dtDataEntregaPropCliente-label").val() + "";
    var dataFinalQuestionamentos = "" + jQuery("#dtDataFinalQuestionamentos-label").val() + "";
    var dataValidadeProposta = "" + jQuery("#dtDataValidadeProposta-label").val() + "";
    var representante = $("#ddlRepresentante").val();
    var cliente = $("#ddlCliente").val();
    var dadosProposta = $("#txtDadosProposta").val();

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

    var quantidade = $("#txtQuantidade").val();

    var arrInstalacao = [];
    $.each($("input[name='checkInstalacao']:checked"), function () {
      arrInstalacao.push($(this).val());
    });

    var arrGarantia = [];
    $.each($("input[name='checkGarantia']:checked"), function () {
      arrGarantia.push($(this).val());
    });

    var arrTipoGarantia = [];
    $.each($("input[name='checkTipoGarantia']:checked"), function () {
      arrTipoGarantia.push($(this).val());
    });

    var arrPrazoGarantia = [];
    $.each($("input[name='checkPrazoGarantia']:checked"), function () {
      arrPrazoGarantia.push($(this).val());
    });

    var arrOutrosServicos = [];
    $.each($("input[name='checkOutrosServicos']:checked"), function () {
      arrOutrosServicos.push($(this).val());
    });

    var arrProduto = Array.prototype.slice.call(document.querySelectorAll('#ddlProduto2 option:checked'), 0).map(function (v, i, a) {
      return v.value;
    });

    _arrAreaId = Array.prototype.slice.call(document.querySelectorAll('#ddlArea2 option:checked'), 0).map(function (v, i, a) {
      return v.value;
    });

    _arrAreaTexto = Array.prototype.slice.call(document.querySelectorAll('#ddlArea2 option:checked'), 0).map(function (v, i, a) {
      return v.text;
    });

    if (tipoAnaliseProposta == "") {
      alert("Escolha o Tipo de Análise!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (sintese == "") {
      alert("Forneça a Síntese!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (identificacaoOportunidade == "") {
      alert("Forneça a Identificação da Oportunidade!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (dataEntregaPropostaCliente == "") {
      alert("Forneça a Data de Entrega da Proposta!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (dataFinalQuestionamentos == "") {
      alert("Forneça a data Final dos Questionamentos!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (dataValidadeProposta == "") {
      alert("Forneça a Data de Validade da Proposta!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;
    }

    if (representante == "0") {
      alert("Escolha o Representante!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;

    }

    if (cliente == "0") {
      alert("Escolha o Cliente!");
      document.getElementById('headingResumoProposta').scrollIntoView();
      return false;

    }

    if (dadosProposta == "") {
      alert("Forneça os Dados da Proposta!");
      document.getElementById('headingDescricaoDetalhada').scrollIntoView();
      return false;
    }

    if (arrSegmento.length == 0) {
      alert("Escolha o Segmento!");
      document.getElementById('headingOportunidade').scrollIntoView();
      return false;
    }

    if (arrSetor.length == 0) {
      alert("Escolha o Setor!");
      document.getElementById('headingOportunidade').scrollIntoView();
      return false;
    }

    if (arrModalidade.length == 0) {
      alert("Escolha a Modalidade!");
      document.getElementById('headingOportunidade').scrollIntoView();
      return false;
    }

    if (quantidade == "") {
      alert("Forneça a Quantidade do Produto!");
      document.getElementById('headingProduto').scrollIntoView();
      return false;
    }

    if (arrInstalacao.length == 0) {
      alert("Escolha a Instalação!");
      document.getElementById('headingProduto').scrollIntoView();
      return false;
    }

    if (arrGarantia.length == 0) {
      alert("Escolha a Garantia!");
      document.getElementById('headingProduto').scrollIntoView();
      return false;
    }

    if (arrPrazoGarantia.length == 0) {
      alert("Escolha a o Prazo de Garantia!");
      document.getElementById('headingProduto').scrollIntoView();
      return false;
    }

    if (arrProduto.length == 0) {
      alert("Escolha o Produto!");
      document.getElementById('headingProduto').scrollIntoView();
      return false;
    }

    /*
    if (_arrAreaTexto.length == 0) {
      alert("Escolha as Áreas Responsáveis!");
      document.getElementById('headingArea').scrollIntoView();
      return false;
    }
    */

    var files = (document.querySelector("#input") as HTMLInputElement).files;

    if (files.length > 0) {

      console.log("files.length", files.length);

      for (var i = 0; i <= files.length - 1; i++) {

        var fsize = files.item(i).size;
        _size = _size + fsize;

        console.log("fsize", fsize);

      }

      if (_size > 15000000) {
        alert("A soma dos arquivos não pode ser maior que 15mega!");
        _size = 0;
        return false;
      }

    }

    //fim valida arquivo

    jQuery("#modalConfirmarIniciarFluxo").modal({ backdrop: 'static', keyboard: false });


  }


  protected salvar() {

    $("#btnSalvar").prop("disabled", true);
    $("#btnIniciarAprovacao").prop("disabled", true);

    if (!_criou) {

      $("#modalConfirmarIniciarFluxo").modal('hide');
      jQuery("#modalCarregando").modal({ backdrop: 'static', keyboard: false });

      //return false;

      console.log("entro no salvar")

      var tipoAnaliseProposta = "";

      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Comercial"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Jurídico"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Representante"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Propostas"; }).prop('selected', true);
      var $options = $('#ddlArea1 option:selected');
      $options.appendTo("#ddlArea2");


      if ($('#checkTipoAnaliseProposta').is(':checked')) { tipoAnaliseProposta = "Proposta" };
      if ($('#checkTipoAnaliseContrato').is(':checked')) { tipoAnaliseProposta = "Contrato" };

      var sintese = $("#txtSintese").val();
      var identificacaoOportunidade = $("#txtIdentificacaoOportunidade").val();

      var dataEntregaPropostaCliente = "" + jQuery("#dtDataEntregaPropCliente-label").val() + "";
      var dataEntregaPropostaClienteDia = dataEntregaPropostaCliente.substring(0, 2);
      var dataEntregaPropostaClienteMes = dataEntregaPropostaCliente.substring(3, 5);
      var dataEntregaPropostaClienteAno = dataEntregaPropostaCliente.substring(6, 10);
      var formDataEntregaPropostaCliente = dataEntregaPropostaClienteAno + "-" + dataEntregaPropostaClienteMes + "-" + dataEntregaPropostaClienteDia;

      var dataFinalQuestionamentos = "" + jQuery("#dtDataFinalQuestionamentos-label").val() + "";
      var dataFinalQuestionamentosDia = dataFinalQuestionamentos.substring(0, 2);
      var dataFinalQuestionamentosMes = dataFinalQuestionamentos.substring(3, 5);
      var dataFinalQuestionamentosAno = dataFinalQuestionamentos.substring(6, 10);
      var formDataFinalQuestionamentos = dataFinalQuestionamentosAno + "-" + dataFinalQuestionamentosMes + "-" + dataFinalQuestionamentosDia;

      var dataValidadeProposta = "" + jQuery("#dtDataValidadeProposta-label").val() + "";
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

      var arrInstalacao = [];
      $.each($("input[name='checkInstalacao']:checked"), function () {
        arrInstalacao.push($(this).val());
      });

      var arrGarantia = [];
      $.each($("input[name='checkGarantia']:checked"), function () {
        arrGarantia.push($(this).val());
      });

      var arrTipoGarantia = [];
      $.each($("input[name='checkTipoGarantia']:checked"), function () {
        arrTipoGarantia.push($(this).val());
      });

      var arrPrazoGarantia = [];
      $.each($("input[name='checkPrazoGarantia']:checked"), function () {
        arrPrazoGarantia.push($(this).val());
      });

      var arrOutrosServicos = [];
      $.each($("input[name='checkOutrosServicos']:checked"), function () {
        arrOutrosServicos.push($(this).val());
      });

      var arrProduto = Array.prototype.slice.call(document.querySelectorAll('#ddlProduto2 option'), 0).map(function (v, i, a) {
        return v.value;
      });

      _arrAreaId = Array.prototype.slice.call(document.querySelectorAll('#ddlArea2 option'), 0).map(function (v, i, a) {
        return v.value;
      });

      _arrAreaTexto = Array.prototype.slice.call(document.querySelectorAll('#ddlArea2 option'), 0).map(function (v, i, a) {
        return v.text;
      });

      console.log("chegou aqui!!!")

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
          Instalacao: arrInstalacao[0],
          Garantia: arrGarantia[0],
          TipoGarantia: arrTipoGarantia[0],
          PrazoGarantia: arrPrazoGarantia[0],
          OutrosServicos: { "results": arrOutrosServicos },
          ProdutoId: { "results": arrProduto }
        })
        .then(response => {

          _idProposta = response.data.ID;

          console.log(_idProposta);

          for (var i = 0; i < _arrAreaId.length; i++) {

            console.log("_arrAreaId[i]", _arrAreaId[i]);
            console.log("_arrAreaTexto[i]", _arrAreaTexto[i]);

            _criou = true;

            _web.lists
              .getByTitle("Tarefas")
              .items.add({
                Title: _arrAreaTexto[i],
                PropostaId: _idProposta,
                DataPlanejadaTermino: formDataEntregaPropostaCliente,
                GrupoSharepointId: _arrAreaId[i]
              })
              .then(response => {

                this.upload();

              }).catch((error: any) => {
                console.log(error);
              });

          }

        }).catch((error: any) => {
          console.log(error);
        });

    }

  }



  protected upload() {

    var files = (document.querySelector("#input") as HTMLInputElement).files;
    var file = files[0];

    //console.log("files.length", files.length);

    if (files.length != 0) {

      _web.lists.getByTitle("AnexosSAP").rootFolder.folders.add(`${_idProposta}`).then(data => {

        for (var i = 0; i < files.length; i++) {

          var nomeArquivo = files[i].name;
          var rplNomeArquivo = nomeArquivo.replace(/[^0123456789.,a-zA-Z]/g, '');

          //alert(rplNomeArquivo);
          //Upload a file to the SharePoint Library
          _web.getFolderByServerRelativeUrl(`${_caminho}/AnexosSAP/${_idProposta}`)
            //.files.add(files[i].name, files[i], true)
            .files.add(rplNomeArquivo, files[i], true)
            .then(function (data) {
              if (i == files.length) {

                console.log("anexou:" + rplNomeArquivo);
                $("#conteudoLoading").modal('hide');
                jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })
                //window.location.href = `home.aspx`;
              }
            });

        }

      }).catch(err => {
        console.log("err", err);
      });

      //const folderAddResult = _web.folders.add(`${_caminho}/Anexos/${_idProposta}`);
      //console.log("foi");

    } else {

      console.log("Gravou!!");
      $("#conteudoLoading").modal('hide');
      jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })

    }

  }




  protected fecharSucesso() {

    $("#modalSucesso").modal('hide');
    window.location.href = `Nova-Proposta-SAP.aspx`;

  }







}
