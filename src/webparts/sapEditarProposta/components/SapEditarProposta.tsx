import * as React from 'react';
import styles from './SapEditarProposta.module.scss';
import { ISapEditarPropostaProps } from './ISapEditarPropostaProps';
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

const divStyle = {
  padding: '0 0 0 20px'
};

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;
var _criou = false;
var _arrAreaId = [];
var _arrAreaTexto = [];
var _caminho;
var _idProposta;
var _size: number = 0;
var _nroAtual: number = 0;
var _nroNovo: number = 0;
var _representante;
var _dataEntregaPropostaCliente;
var _datadataFinalQuestionamentos;
var _dataValidadeProposta;
var _cliente;
var _dadosProposta;
var _arrSegmento = [];
var _arrSetor = [];
var _arrModalidade = [];
var _instalacao = [];
var _garantia = [];
var _tipoGarantia = [];
var _prazoGarantia = [];
var _arrOutrosServicos = [];
var _arrProduto = [];
var _arrAreas = [];
var _serverRelativeUrl;
var _nomeArquivo;
var _elemento;
var _elemento2;
var _siteurl;
var _arrAreasAntiga = [];
var _txtCliente;
var _areaAnexo;
var _pastaCriada;
var _siteAntigo;
var _idAntigo;
var _numeroProposta;


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
  itemsAreasAnexos: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsResponsavelProposta: [
    {
      "ID": "",
      "Title": "",
      "Responsavel": { "Title": "" }
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
  itemsDataEntregaPropostaCliente: any;
  valorItemsRepresentante: "",
  valorItemsCliente: "",
  valorItemsResponsavelProposta: any;
  valorCheckedSegmento;



}

export default class SapEditarProposta extends React.Component<ISapEditarPropostaProps, IReactGetItemsState> {

  public constructor(props: ISapEditarPropostaProps, state: IReactGetItemsState) {
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
      itemsAreasAnexos: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsResponsavelProposta: [
        {
          "ID": "",
          "Title": "",
          "Responsavel": { "Title": "" }
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
      itemsDataEntregaPropostaCliente: "",
      valorItemsRepresentante: "",
      valorItemsCliente: "",
      valorItemsResponsavelProposta: "",
      valorCheckedSegmento: "any"


    };


  }

  public componentDidMount() {


    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    console.log("_caminho", _caminho);

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idProposta = parseInt(queryParms.getValue("PropostasID"));

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

    document
      .getElementById("btExcluirAnexo")
      .addEventListener("click", (e: Event) => this.excluirAnexo());

    document
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());


    $("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/Images1/loading.gif'/>
    <br/>Aguarde....<br/><br/>
    Dependendo do tamanho do anexo e a velocidade<br>
     da Internet essa ação pode demorar um pouco. <br>
     Não fechar a janela!<br/><br/>`);


    this.handler();


  }





  public render(): React.ReactElement<ISapEditarPropostaProps> {




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
                        <input type="text" className="form-control" id="txtIdentificacaoOportunidade" />
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataEntregaPropostaCliente">Data da entrega da Proposta ao Cliente</label><span className="required"> *</span>
                        <DatePicker minDate={this.addDaysWRONG()} value={_dataEntregaPropostaCliente} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="datePicker" id='dtDataEntregaPropCliente' />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataFinalQuestionamentos">Data final de questionamentos</label>
                        <DatePicker minDate={new Date()} value={_datadataFinalQuestionamentos} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="datePicker" id='dtDataFinalQuestionamentos' />
                      </div>
                      <div className="form-group col-md-4">
                        <label htmlFor="dtDataValidadeProposta">Data de validade da Proposta</label>
                        <DatePicker minDate={new Date()} value={_dataValidadeProposta} formatDate={this.onFormatDate} isMonthPickerVisible={false} className="datePicker" id='dtDataValidadeProposta' />
                      </div>
                    </div>
                  </div>

                  <div className="form-group">
                    <div className="form-row">
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlRepresentante">Representante</label><span className="required"> *</span>
                        <select id="ddlRepresentante" className="form-control" value={this.state.valorItemsRepresentante} disabled  >
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
                        <select id="ddlCliente" className="form-control" value={this.state.valorItemsCliente} onChange={(e) => this.onChangeCliente(e.target.value)} >
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
                      <div className="form-group col-md-6">
                        <label htmlFor="ddlResponsavelProposta">Responsável da Proposta</label><span className="required"> *</span>
                        <select id="ddlResponsavelProposta" className="form-control" value={this.state.valorItemsResponsavelProposta} onChange={(e) => this.onChangeResponsavelProposta(e.target.value)}>
                          <option value="0" selected>Selecione...</option>
                          {this.state.itemsResponsavelProposta.map(function (item, key) {
                            return (
                              <option value={item.Responsavel.Title}>{item.Responsavel.Title}</option>
                            );
                          })}
                        </select>
                      </div>
                      <div className="form-group col-md-3">
                        <label htmlFor="txtPropostaRevisadaReferencia">Proposta revisada/referência</label>
                        <input type="text" className="form-control" id="txtPropostaRevisadaReferencia" />
                      </div>
                      <div className="form-group col-md-3">
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
                    <RichText className="editorRichTex" value={_dadosProposta}
                      onChange={(text) => this.onTextChange(text)}
                    />
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
                        {this.state.itemsSegmento.map((item) => {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkSegmento' defaultChecked={_arrSegmento.indexOf(item) !== -1} type="checkbox" value={item} />
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

                          var checado = false;
                          if (_arrSetor == item) checado = true;

                          return (

                            <div className="form-check">
                              <input className="form-check-input" type="radio" defaultChecked={checado} name="checkSetor" value={item} />
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

                          var checado = false;
                          if (_arrModalidade == item) checado = true;

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkModalidade' type="radio" defaultChecked={checado} value={item} />
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

                          var checado = false;
                          if (_instalacao == item) checado = true;

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkInstalacao' type="radio" defaultChecked={checado} value={item} />
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

                          var checado = false;
                          if (_garantia == item) checado = true;

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkGarantia' type="radio" defaultChecked={checado} value={item} />
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

                          var checado = false;
                          if (_tipoGarantia == item) checado = true;

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkTipoGarantia' type="radio" defaultChecked={checado} value={item} />
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

                        <input type="number" style={{ "width": "120px" }} className="form-control" id="txtPrazoGarantia" />

                      </div>
                      <div className="form-group col-md-2">
                        <label htmlFor="txtTitulo">Outros serviços</label><br></br>

                        {this.state.itemsOutrosServicos.map(function (item, key) {

                          return (

                            <div className="form-check">
                              <input className="form-check-input" name='checkOutrosServicos' type="checkbox" defaultChecked={_arrOutrosServicos.indexOf(item) !== -1} value={item} />
                              <label className="form-check-label">
                                {item}
                              </label>
                            </div>

                          );
                        })}

                      </div>
                      <div className="form-group col-md-8">
                        <label htmlFor="ddlProduto">Produto</label><span className="required"> *</span>
                        <table>
                          <tr>
                            <td>
                              <div className="col-sm-2">
                                <select multiple={true} id='ddlProduto1' className="form-control" name="ddlProduto1" style={{ "height": "194px", "width": "200px" }}>

                                  {this.state.itemsProdutos.map(function (item, key) {

                                    if (_arrProduto.indexOf(item.ID) == -1) {
                                      return (
                                        <option className="optProduto" value={item.ID}>{item.Title}</option>
                                      );

                                    }

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
                                  {this.state.itemsProdutos.map(function (item, key) {

                                    if (_arrProduto.indexOf(item.ID) !== -1) {
                                      return (
                                        <option className="optProduto" value={item.ID}>{item.Title}</option>
                                      );

                                    }

                                  })}
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

                              if (_arrAreas.indexOf(item.ID) == -1) {
                                return (
                                  <option className="optArea" value={item.ID}>{item.Title}</option>
                                );

                              }

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

                            {this.state.itemsAreas.map(function (item, key) {

                              if (_arrAreas.indexOf(item.ID) != -1) {
                                return (
                                  <option className="optArea" value={item.ID}>{item.Title}</option>
                                );

                              }

                            })}
                          </select>
                        </div>
                      </td>
                    </tr>
                  </table>
                  <br></br>
                  <p className="text-info">&nbsp;&nbsp;&nbsp; As seguintes áreas já são adicionadas a Proposta automaticamente: Comercial, Jurídico, Representante, Propostas </p>
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
                          <th scope="col">Ação</th>
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
                            {this.state.itemsAreasAnexos.map(function (item, key) {
                              return (
                                <option value={item.ID}>{item.Title}</option>
                              );
                            })}
                          </select>
                        </div>
                        <div className="form-group col-md" >
                        </div>
                      </div>
                    </div>

                    <p className='text-info'>Total máximo permitido: 15 MB</p>


                  </div>

                </div>

              </div>
            </div>

            <br></br>

            <span className='text-info'>Criador por: <span id='txtCriadoPor'></span> no dia <span id='txtCriadoData'></span> às <span id='txtCriadoHora'>11:30</span></span><br></br>
            <span className='text-info'>Modificado por: <span id='txtModificadoPor'></span> no dia <span id='txtModificadoData'></span> às <span id='txtModificadoHora'>11:30</span></span>


            <br></br>

            <div className="text-right">
              <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
              <button style={{ "margin": "2px" }} id="btnIniciarAprovacao" className="btn btn-success" >Salvar</button>
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
                Deseja realmente salvar as alterações?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btIniciarFluxo" type="button" className="btn btn-primary">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalConfirmarExcluirAnexo" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">
                Deseja realmente excluir o arquivo?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btExcluirAnexo" type="button" className="btn btn-primary">Excluir</button>
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
                Proposta atualizada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucesso" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoAnexoExcluido" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Anexo excluido com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">OK</button>
              </div>
            </div>
          </div>
        </div>

      </>


    );
  }


  protected changeSegmento = (val) => {

    this.setState({
      valorCheckedSegmento: true
    });

    return {
      value: 'select'
    }

    //console.log("elem",elem);

    this.setState({
      // itemsSegmento: val
    });


  }

  private onTextChange = (newText: string) => {
    _dadosProposta = newText;
    return newText;
  }

  private onChangeCliente = (val) => {
    this.setState({
      valorItemsCliente: val,
    });
  }

  private onChangeResponsavelProposta = (val) => {
    this.setState({
      valorItemsResponsavelProposta: val,
    });
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
    var reactHandlerResponsavelProposta = this;
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
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerRepresentante.setState({
          itemsRepresentante: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Responsavel Proposta')/items?$top=4999&$select=ID,Responsavel/Title&$expand=Responsavel`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log("resultDataResponsavel", resultData);
        reactHandlerResponsavelProposta.setState({

          itemsResponsavelProposta: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Clientes')/items?$top=4999&$filter=Ativo eq 1&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerClientes.setState({
          itemsClientes: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
      }
    });

    /*
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
        console.log(jqXHR.responseText);
      }
    });
    */

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
        console.log(jqXHR.responseText);
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
        console.log(jqXHR.responseText);
      }
    });


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Areas')/items?$top=4999&$filter=((Ativo eq 1) and (Title ne 'Comercial') and (Title ne 'Jurídico') and (Title ne 'Representante') and (Title ne 'Propostas'))&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerAreas.setState({
          itemsAreas: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Areas')/items?$top=4999&$filter=Ativo eq 1&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerAreas.setState({
          itemsAreasAnexos: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });



    this.getProposta();
    this.getTarefas();
    this.getAnexos();

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


  protected getProposta() {

    console.log("entrou no proposta");

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('PropostasSAP')/items?$select=ID,Title,TipoAnalise,Status,Numero,IdentificacaoOportunidade,DataEntregaPropostaCliente,DataFinalQuestionamentos,DataValidadeProposta,Representante/ID,Cliente/ID,PropostaRevisadaReferencia,CondicoesPagamento,DadosProposta,Segmento,Setor,Modalidade,NumeroEditalRFPRFQRFI,Instalacao,Quantidade,Garantia,TipoGarantia,PrazoGarantia,OutrosServicos,Produto/ID,ResponsavelProposta,Created,Author/Title,Modified,Editor/Title,IDAntigo,SiteAntigo,PastaCriada&$expand=Representante,Cliente,Produto,Author,Editor&$filter=ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        console.log("resultData Proposta", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            _siteAntigo = resultData.d.results[i].SiteAntigo;
            _pastaCriada = resultData.d.results[i].PastaCriada;
            _idAntigo = resultData.d.results[i].IDAntigo;
            _numeroProposta = resultData.d.results[i].Numero;

            var tipoAnalise = resultData.d.results[i].TipoAnalise;

            console.log("resultData.d.results[i].DataEntregaPropostaCliente", resultData.d.results[i].DataEntregaPropostaCliente);

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

            var dataEntregaPropostaCliente = resultData.d.results[i].DataEntregaPropostaCliente;
            var dataFinalQuestionamentos = resultData.d.results[i].DataFinalQuestionamentos;
            var dataValidadeProposta = resultData.d.results[i].DataValidadeProposta;

            var status = resultData.d.results[i].Status;

            console.log("status", status);

            if (status != "Em análise") {

              $("#btnIniciarAprovacao").prop("disabled", true);

            }

            if (dataEntregaPropostaCliente != null) {

              var dtDataEntregaPropostaCliente = new Date(dataEntregaPropostaCliente);
              _dataEntregaPropostaCliente = dtDataEntregaPropostaCliente;

            } else _dataEntregaPropostaCliente = null;


            if (dataFinalQuestionamentos != null) {

              var dtdataFinalQuestionamentos = new Date(dataFinalQuestionamentos);
              _datadataFinalQuestionamentos = dtdataFinalQuestionamentos;

            } else _datadataFinalQuestionamentos = null;


            if (dataValidadeProposta != null) {

              var dtDataValidadeProposta = new Date(dataValidadeProposta);
              _dataValidadeProposta = dtDataValidadeProposta;

            } else _dataValidadeProposta = null;


            if (tipoAnalise == "Proposta") jQuery("#checkTipoAnaliseProposta").attr('checked', 'true');
            else if (tipoAnalise == "Contrato") jQuery("#checkTipoAnaliseContrato").attr('checked', 'true');

            console.log("resultData.d.results[i].Representante.ID", resultData.d.results[i].Representante.ID);

            jQuery("#txtSintese").val(resultData.d.results[i].Title);
            jQuery("#txtIdentificacaoOportunidade").val(resultData.d.results[i].IdentificacaoOportunidade);

            console.log("resultData.d.results[i].Cliente.ID", resultData.d.results[i].Cliente.ID)

            this.setState({
              valorItemsRepresentante: resultData.d.results[i].Representante.ID,
              valorItemsCliente: resultData.d.results[i].Cliente.ID,
            });

            var itemsResponsavelProposta = resultData.d.results[i].ResponsavelProposta;

            if (itemsResponsavelProposta == null) {

              this.setState({
                valorItemsResponsavelProposta: 0
              });

            } else {

              this.setState({
                valorItemsResponsavelProposta: resultData.d.results[i].ResponsavelProposta
              });


            }



            //_representante = resultData.d.results[i].Representante.ID;
            //_cliente = resultData.d.results[i].Cliente.ID;
            //_responsavelProposta = resultData.d.results[i].ResponsavelProposta;

            _dadosProposta = resultData.d.results[i].DadosProposta;

            jQuery("#txtPropostaRevisadaReferencia").val(resultData.d.results[i].PropostaRevisadaReferencia);
            jQuery("#txtCondicoesPagamento").val(resultData.d.results[i].CondicoesPagamento);
            jQuery("#txtDadosProposta").val(resultData.d.results[i].DadosProposta);

            _arrSegmento = resultData.d.results[i].Segmento.results;
            _arrSetor = resultData.d.results[i].Setor;
            _arrModalidade = resultData.d.results[i].Modalidade;

            console.log("_arrSetor1", _arrSetor);

            jQuery("#txtNumeroEditalRFPRFQRFI").val(resultData.d.results[i].NumeroEditalRFPRFQRFI);
            jQuery("#txtQuantidade").val(resultData.d.results[i].Quantidade);

            _instalacao = resultData.d.results[i].Instalacao;
            _garantia = resultData.d.results[i].Garantia;
            _tipoGarantia = resultData.d.results[i].TipoGarantia;
            _prazoGarantia = resultData.d.results[i].PrazoGarantia;

            jQuery("#txtPrazoGarantia").val(_prazoGarantia);

            var outrosServicos = resultData.d.results[i].OutrosServicos;

            if (outrosServicos != null) {

              _arrOutrosServicos = resultData.d.results[i].OutrosServicos.results;

            }

            var arrProduto = [];
            arrProduto = resultData.d.results[i].Produto.results;

            var tamArrProduto = resultData.d.results[i].Produto.results.length;

            for (i = 0; i < tamArrProduto; i++) {

              _arrProduto.push(arrProduto[i].ID);

            }






          }

        }

        //console.log("_arrProdutoZ", _arrProduto);

      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }



    })

  }


  protected getTarefas() {

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$select=ID,Title,GrupoSharepoint/ID,GrupoSharepoint/Title,Status,DataPlanejadaTermino,Modified,DataRealTermino,Justificativa&$expand=GrupoSharepoint&$filter=(AntigoPropostaNumero eq ` + _numeroProposta + `.000000000) and (Status ne 'Em análise')`,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: async (resultData) => {

        console.log("resultData.d.results tarefas antigas", resultData.d.results);

        //var reactHandlerAreaAnexo = this;

        //reactHandlerAreaAnexo.setState({
         // itemsAreasAnexos: resultData.d.results
        //});

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var titulo = resultData.d.results[i].Title;
            var rplTitulo = titulo.replaceAll("Avaliação da Área (", "");
            var rplTitulo = rplTitulo.replaceAll(")", "");

            _arrAreas.push(resultData.d.results[i].GrupoSharepoint.ID);
            _arrAreasAntiga.push(rplTitulo);

          }

        }


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });



    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$select=ID,Title,GrupoSharepoint/ID&$expand=GrupoSharepoint&$orderby=Title&$filter=Proposta/ID eq ` + _idProposta,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        console.log("resultData Proposta", resultData);

        //var reactHandlerAreaAnexo = this;

        //reactHandlerAreaAnexo.setState({
         // itemsAreasAnexos: resultData.d.results
       // });

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            //_arrSegmento = resultData.d.results[i].Title;

            if (_arrAreas.indexOf(resultData.d.results[i].GrupoSharepoint.ID) == -1) {
            _arrAreas.push(resultData.d.results[i].GrupoSharepoint.ID);
            }

            if (_arrAreasAntiga.indexOf(resultData.d.results[i].Title) == -1) {
              _arrAreasAntiga.push(resultData.d.results[i].Title);
            }

          }

        }

        //console.log("_arrAreas", _arrAreas);
        console.log("_arrAreasAntiga", _arrAreasAntiga);


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }



    })


  }

  /*
  protected getAnexosOld() {

    //get anexos da biblioteca

    var montaAnexo = "";

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Propostas-SAP-Editar.aspx", "");

    //var relative = "/sites/bit-hml";
    var idItem = 0;

    console.log("_bitNumero", _idProposta);

    console.log("caminho", `${strRelativeURL}/AnexosSAP/${_idProposta}`);


    _web.getFolderByServerRelativeUrl(`${strRelativeURL}/AnexosSAP/${_idProposta}`)
      .expand("Folders, Files, ListItemAllFields").get().then(r => {
        console.log("r", r);
        r.Folders.forEach(item => {
          console.log("item-doc", item);
          console.log("entrou em folder");
        })
        r.Files.forEach(item => {
          console.log("entrou em files");

          console.log("item", item);
          idItem++;
          $("#conteudoAnexoNaoEncontrado").hide();
          montaAnexo = `<a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a>&nbsp;<a id="btnExcluirAnexo${idItem}" style="cursor:pointer" >Excluir</a> <br/>`

          $("#conteudoAnexo").append(montaAnexo);

        })

      }).catch((error: any) => {
        console.log("Erro Anexo da biblioteca: ", error);
      });


    //fim anexos da biblioteca


  }

  */

  protected getAnexos() {

    var montaAnexo = "";
    var montaAnexo2 = "";
    var relativeURL = window.location.pathname;
    var strRelativeURL = relativeURL.replace("SitePages/Propostas-SAP-Editar.aspx", "");
    var idItem = 0;

    if (_siteAntigo == "Sim") {

      jquery.ajax({
        url: `${this.props.siteurl}/_api/web/lists/getbytitle('Anexos')/items?$select=Title,AreaSelecionada,Created,Author/Title,File/ServerRelativeUrl&$expand=Author,File&$filter=Proposta/Numero eq ` + _numeroProposta,
        type: "GET",
        async: false,
        headers: { 'Accept': 'application/json; odata=verbose;' },
        success: function (resultData) {

          console.log("resultData anexos antigos", resultData);

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

    _web.getFolderByServerRelativeUrl(`${strRelativeURL}/AnexosSAP/${_idProposta}`).files
      .expand('ListItemAllFields', 'Author').get().then(r => {

        console.log("r1", r);

        if (r.length != 0) {

          r.forEach(item => {

            idItem++;

            var criado = new Date(item.TimeCreated);
            var criadoData = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear();
            var criadoHora = criado.getHours() + ":" + ("0" + (criado.getMinutes() + 1)).slice(-2) + ":" + criado.getSeconds();

            montaAnexo = `<tr id="anexo${idItem}"> 
          <td style="word-break: break-word;"><a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a></td>
          <td style="word-break: break-word;">${item.ListItemAllFields.Area}</td>
          <td style="word-break: break-word;">${criadoData} ${criadoHora}</td>
          <td style="word-break: break-word;">${item.Author.Title}</td>
          <td style="word-break: break-word; width: 100px;"><button id="btnExcluirAnexo${idItem}" type="button" class="btn btn-secondary btn-sm">Excluir</button></td>
          </tr> 
          `

            jQuery("#conteudoAnexo").append(montaAnexo);

            document
              .getElementById(`btnExcluirAnexo${idItem}`)
              .addEventListener("click", (e: Event) => this.confirmarExcluirAnexo(item.ServerRelativeUrl, item.Name, `anexo${idItem}`, `btnExcluirAnexo${idItem}`));

          });



        }

      }).catch((error: any) => {
        console.log("Erro onChangeCliente: ", error);
      });




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
    var responsavelProposta = $("#ddlResponsavelProposta").val();
    _representante = representante;
    var cliente = $("#ddlCliente").val();
    var dadosProposta = $("#txtDadosProposta").val();
    var prazoGarantia = $("#txtPrazoGarantia").val();

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

    if (responsavelProposta == "0") {
      alert("Escolha o Responsável pela Proposta!");
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

    if (prazoGarantia == "") {
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

      _areaAnexo = $("#ddlAreaAnexo option:selected").text();

      if (_areaAnexo == "Selecione...") {
        alert("Escolha a Área do anexo");
        return false;
      }


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

  protected confirmarExcluirAnexo(serverRelativeUrl, nomeArquivo, elemento, elemento2) {

    _serverRelativeUrl = serverRelativeUrl;
    _nomeArquivo = nomeArquivo;
    _elemento = elemento;
    _elemento2 = elemento2;

    console.log("_nomeArquivo", _nomeArquivo);
    console.log("_elemento", _elemento);
    console.log("_elemento2", _elemento2);

    //return false;

    jQuery("#modalConfirmarExcluirAnexo").modal({ backdrop: 'static', keyboard: false });

  }

  protected async excluirAnexo() {

    console.log("_serverRelativeUrl", _serverRelativeUrl);
    console.log("_nomeArquivo", _nomeArquivo);
    console.log("_elemento", _elemento);
    console.log("_elemento2", _elemento2);

    const list = _web.lists.getByTitle("AnexosSAP");

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Propostas-SAP-Editar.aspx", "");

    _web.getFolderByServerRelativePath(`${strRelativeURL}/AnexosSAP/${_idProposta}`).files.getByName(_nomeArquivo).delete()

      .then(async response => {

        $("#modalConfirmarExcluirAnexo").modal('hide');

        jQuery(`#${_elemento}`).hide();
        jQuery(`#${_elemento2}`).hide();
        jQuery("#modalSucessoAnexoExcluido").modal({ backdrop: 'static', keyboard: false });

      }).catch((error: any) => {
        console.log("Erro em excluirAnexo " + error);

      })

  }

  protected async salvar() {

    $("#btnSalvar").prop("disabled", true);
    $("#btnIniciarAprovacao").prop("disabled", true);

    if (!_criou) {

      $("#modalConfirmarIniciarFluxo").modal('hide');
      jQuery("#modalCarregando").modal({ backdrop: 'static', keyboard: false });

      //return false;

      console.log("entro no salvar")

      var tipoAnaliseProposta = "";

      /*

      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Comercial"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Jurídico"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Representante"; }).prop('selected', true);
      jQuery('#ddlArea1 option').filter(function () { return $(this).html() == "Propostas"; }).prop('selected', true);
      var $options = $('#ddlArea1 option:selected');
      $options.appendTo("#ddlArea2");

      */


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

      if (dataFinalQuestionamentos != "") {
        var dataFinalQuestionamentosDia = dataFinalQuestionamentos.substring(0, 2);
        var dataFinalQuestionamentosMes = dataFinalQuestionamentos.substring(3, 5);
        var dataFinalQuestionamentosAno = dataFinalQuestionamentos.substring(6, 10);
        var formDataFinalQuestionamentos = dataFinalQuestionamentosAno + "-" + dataFinalQuestionamentosMes + "-" + dataFinalQuestionamentosDia;
      }
      else formDataFinalQuestionamentos = null;

      var dataValidadeProposta = "" + jQuery("#dtDataValidadeProposta-label").val() + "";

      if (dataValidadeProposta != "") {
        var dataValidadePropostaDia = dataValidadeProposta.substring(0, 2);
        var dataValidadePropostaMes = dataValidadeProposta.substring(3, 5);
        var dataValidadePropostaAno = dataValidadeProposta.substring(6, 10);
        var formDataValidadeProposta = dataValidadePropostaAno + "-" + dataValidadePropostaMes + "-" + dataValidadePropostaDia;
      }
      else formDataValidadeProposta = null;

      console.log("sintese", sintese);
      console.log("tipoAnaliseProposta", tipoAnaliseProposta);
      console.log("identificacaoOportunidade", identificacaoOportunidade);
      console.log("formDataEntregaPropostaCliente", formDataEntregaPropostaCliente);
      console.log("formDataFinalQuestionamentos", formDataFinalQuestionamentos);
      console.log("formDataValidadeProposta", formDataValidadeProposta);

      var responsavelProposta = $("#ddlResponsavelProposta").val();
      var cliente = $("#ddlCliente").val();
      _txtCliente = $('#ddlCliente :selected').text();
      var propostaRevisadaReferencia = $("#txtPropostaRevisadaReferencia").val();
      var SST = $("#txtSST").val();
      var condicoesPagamento = $("#txtCondicoesPagamento").val();
      var justificativaFinal = $("#txtJustificativa").val();
      var prazoGarantia = $("#txtPrazoGarantia").val();


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

      /*
            jquery.ajax({
      
              url: `${this.props.siteurl}/_api/web/lists/getbytitle('Areas')/items?$select=ID,Title&$filter=(Title eq 'Comercial') or (Title eq 'Jurídico') or (Title eq 'Representante') or (Title eq 'Propostas')`,
              type: "GET",
              async: false,
              headers: { 'Accept': 'application/json; odata=verbose;' },
              success: function (resultData) {
                if (resultData.d.results.length > 0) {
                  for (var i = 0; i < resultData.d.results.length; i++) {
                    _arrAreaId.push(resultData.d.results[i].ID);
                    _arrAreaTexto.push(resultData.d.results[i].Title);
                  }
                }
              },
              error: function (jqXHR, textStatus, errorThrown) {
                console.log(textStatus);
              }
            });
      
      */

      await _web.lists
        .getByTitle("PropostasSAP")
        .items.getById(_idProposta).update({
          Title: sintese,
          TipoAnalise: tipoAnaliseProposta,
          IdentificacaoOportunidade: identificacaoOportunidade,
          DataEntregaPropostaCliente: formDataEntregaPropostaCliente,
          DataFinalQuestionamentos: formDataFinalQuestionamentos,
          DataValidadeProposta: formDataValidadeProposta,
          //RepresentanteId: representante,
          ClienteId: cliente,
          ResponsavelProposta: responsavelProposta,
          PropostaRevisadaReferencia: propostaRevisadaReferencia,
          SST: SST,
          CondicoesPagamento: condicoesPagamento,
          DadosProposta: _dadosProposta,
          JustificativaFinal: justificativaFinal,
          Segmento: { "results": arrSegmento },
          Setor: arrSetor[0],
          Modalidade: arrModalidade[0],
          NumeroEditalRFPRFQRFI: numeroEditalRFPRFQRFI,
          Quantidade: quantidade,
          Instalacao: arrInstalacao[0],
          Garantia: arrGarantia[0],
          TipoGarantia: arrTipoGarantia[0],
          PrazoGarantia: prazoGarantia,
          OutrosServicos: { "results": arrOutrosServicos },
          ProdutoId: { "results": arrProduto }
        })
        .then(async response => {



          if (_arrAreaId.length != 0) {


            for (var x = 0; x < _arrAreaId.length; x++) {

              console.log("_arrAreaId[x]", _arrAreaId[x]);
              console.log("_arrAreaTexto[x]", _arrAreaTexto[x]);

              _criou = true;

              console.log("_arrAreasAntiga.indexOf(_arrAreaTexto[x])", _arrAreasAntiga.indexOf(_arrAreaTexto[x]));


              if (_arrAreasAntiga.indexOf(_arrAreaTexto[x]) == -1) {

                await _web.lists
                  .getByTitle("Tarefas")
                  .items.add({
                    Title: _arrAreaTexto[x],
                    PropostaId: _idProposta,
                    DataPlanejadaTermino: formDataEntregaPropostaCliente,
                    GrupoSharepointId: _arrAreaId[x],
                    Cliente: _txtCliente
                  })
                  .then(response => {

                    var last = (_arrAreaId.length) - 1;
                    console.log("last", last);
                    console.log("x", x);
                    if (x == last) this.upload();

                  }).catch((error: any) => {
                    console.log(error);
                  });

              } else {

                var last = (_arrAreaId.length) - 1;
                console.log("last", last);
                console.log("x", x);
                if (x == last) this.upload();

              }

            }

          } else {
            this.upload();
          }

        })

    }

  }

  protected async upload() {

    var files = (document.querySelector("#input") as HTMLInputElement).files;
    var file = files[0];

    console.log("files.length", files.length);

    if (files.length != 0) {

      if (_pastaCriada != "Sim") {

        _web.lists.getByTitle("AnexosSAP").rootFolder.folders.add(`${_idProposta}`);

        await _web.lists
        .getByTitle("PropostasSAP")
        .items.getById(_idProposta).update({
          PastaCriada: "Sim",
        })
        .then(async response => {


        }).catch(err => {
          console.log("err", err);
        });
        
      }

      for (var i = 0; i < files.length; i++) {

        var nomeArquivo = files[i].name;
        var rplNomeArquivo = nomeArquivo.replace(/[^0123456789.,a-zA-Z]/g, '');

        //alert(rplNomeArquivo);
        //Upload a file to the SharePoint Library
        _web.getFolderByServerRelativeUrl(`${_caminho}/AnexosSAP/${_idProposta}`)
          //.files.add(files[i].name, files[i], true)
          .files.add(rplNomeArquivo, files[i], true)
          .then(function (data) {

            data.file.getItem().then(async item => {
              var idAnexo = item.ID;

              await _web.lists
                .getByTitle("AnexosSAP")
                .items.getById(idAnexo).update({
                  Area: _areaAnexo,
                })
                .then(async response => {

                  if (i == files.length) {
                    console.log("anexou:" + rplNomeArquivo);
                    $("#conteudoLoading").modal('hide');
                    jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })
                  }
                }).catch(err => {
                  console.log("err", err);
                });

            })

          });

      }

    } else {

      console.log("Gravou!!");
      $("#conteudoLoading").modal('hide');
      jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })

    }

  }

  protected fecharSucesso() {

    $("#modalSucesso").modal('hide');
    window.location.href = `Propostas.aspx`;

  }

  protected voltar() {

    history.back();

  }






}
