import * as React from 'react';
import styles from './BitNovoItem.module.scss';
import { IBitNovoItemProps } from './IBitNovoItemProps';
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

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");

var testeGus;
var _addUsersAprovadorEngenharia = [];
var _addUsersAprovadorGeral = [];
var _addUsersDestinatariosAdicionais = [];
var _msgValidacao;
var _userTitle;
var _idBit;
var _caminho;
var _web;

export interface IReactGetItemsState {
  items: [
    {
      "ID": "",
      "Title": "",
      "Ativo": "",
    }],
  itemsCliente: [
    {
      "ID": "",
      "Title": "",
      "Ativo": "",
    }],
  itemsAplicacoes: [
    {
      "ID": "",
      "Title": "",
      "Ativo": "",
    }],
  itemsJustificativa: [
    {
      "ID": "",
      "Title": "",
      "Ativo": "",
    }],
  showmessageBar: boolean; //to show/hide message bar on success
  message: string; // what message to be displayed in message bar
  itemID: number; // current item ID after create new item is clicked
  addUsersAprovadorEngenharia: [];
  addUsersAprovadorGeral: [];
  addUsersDestinatariosAdicionais: [];
  defaultmyusers: [];
}


export default class BitNovoItem extends React.Component<IBitNovoItemProps, IReactGetItemsState> {


  public constructor(props: IBitNovoItemProps, state: IReactGetItemsState) {
    super(props);
    this.state = {
      items: [
        {
          "ID": "",
          "Title": "",
          "Ativo": "",
        }
      ],
      itemsCliente: [
        {
          "ID": "",
          "Title": "",
          "Ativo": "",
        }
      ],
      itemsAplicacoes: [
        {
          "ID": "",
          "Title": "",
          "Ativo": "",
        }],
      itemsJustificativa: [
        {
          "ID": "",
          "Title": "",
          "Ativo": "",
        }],
      showmessageBar: false,
      message: "",
      itemID: 0,
      addUsersAprovadorEngenharia: [],
      addUsersAprovadorGeral: [],
      addUsersDestinatariosAdicionais: [],
      defaultmyusers: []
    };
  }


  public async componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    $("#divJustificativaBit").hide();

    _caminho = this.props.siteurl;

    console.log("_caminho", _caminho);

    sp.web.currentUser.get().then(f => {
      _userTitle = f.Title;
      console.log("_userTitle2", _userTitle);
    })


    var reactHandler = this;

    console.log(`${this.props.siteurl}/_api/web/lists/getbytitle('Produtos')/items?$top=4999&$filter=Ativo eq 1&$orderby= ID desc`);

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Produtos')/items?$top=4999&$filter=Ativo eq 1&$orderby= ID desc`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandler.setState({
          items: resultData.d.results
        });
        console.log("resultData.d.results1", resultData.d.results);
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    var reactHandler2 = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Clientes')/items?$top=4999&$filter=Ativo eq 1&$orderby= ID desc`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandler2.setState({
          itemsCliente: resultData.d.results
        });
        console.log("resultData.d.results2", resultData.d.results)
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    var reactHandler3 = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Aplicações')/items?$top=4999&$filter=Ativo eq 1&$orderby= ID desc`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log("entrouuuu");
        reactHandler3.setState({
          itemsAplicacoes: resultData.d.results
        });
        console.log("resultData.d.results3", resultData.d.results)
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    var reactHandler4 = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Justificativa Cancelar BIT')/items?$filter=Ativo eq 1&$orderby= ID desc`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log("entrouuuu");
        reactHandler4.setState({
          itemsJustificativa: resultData.d.results
        });
        console.log("resultData.d.results4", resultData.d.results)
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


  }


  public render(): React.ReactElement<IBitNovoItemProps> {


    return (


      <><div className="container-fluid border">

        <div className="form-group">
          <div className="form-row">
            <div className="form-group col-md-8">
              <label htmlFor="txtTitulo">Título</label><span className={styles.required}> *</span>
              <input type="text" className="form-control" id="txtTitulo" />
            </div>
            <div className="form-group col-md-4">
              <label htmlFor="ddlOrigemBit">Origem BIT</label><span className={styles.required}> *</span>
              <select id="ddlOrigemBit" className="form-control">
                <option value="Engenharia AT" selected>Engenharia AT</option>
              </select>
            </div>
          </div>
        </div>


        <div className="form-row">
          <div className="form-group col-md-12">
            <label htmlFor="produto1">Produto</label><span className={styles.required}> *</span>
            <table>
              <tr>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="produtos1" className="form-control" name="produto1" style={{ "height": "194px", "width": "300px" }}>
                      {this.state.items.map((item) => (
                        <option className="optProduto" value={item.ID}>{item.Title}</option>
                      ))}
                    </select>
                  </div>
                </td>
                <td>
                  <div>
                    <input type="button" className="btn btn-light" id="addButtonProduto" onClick={this.addButtonProduto} value="Adicionar >" alt="Salvar" /></div><br />
                  <input type="button" className="btn btn-light" id="removeButtonProduto" onClick={this.removeButtonProduto} value="< Remover"
                    alt="Salvar" />
                </td>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="produtos2" className="form-control" name="produto2" style={{ "height": "194px", "width": "300px" }}>
                    </select>
                  </div>
                </td>
              </tr>
            </table>
          </div>
        </div>

        <div className="form-row">
          <div className="form-group col-md-12">
            <label htmlFor="cliente1">Cliente</label><span className={styles.required}> *</span>
            <table>
              <tr>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="cliente1" className="form-control" name="cliente1"
                      style={{ "height": "194px", "width": "300px" }}>
                      {this.state.itemsCliente.map(function (item, key) {
                        return (
                          <option className="optCliente" value={item.ID}>{item.Title}</option>
                        );
                      })}

                    </select>
                  </div>
                </td>
                <td>
                  <div>
                    <input type="button" className="btn btn-light" id="addButtonCliente" onClick={this.addButtonCliente} value="Adicionar >" alt="Salvar" /></div><br />
                  <input type="button" className="btn btn-light" id="removeButtonCliente" onClick={this.removeButtonCliente} value="< Remover"
                    alt="Salvar" />
                </td>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="cliente2" className="form-control" name="cliente2"
                      style={{ "height": "194px", "width": "300px" }}>
                    </select>
                  </div>
                </td>
              </tr>
            </table>

          </div>
        </div>


        <div className="form-row">
          <div className="form-group col-md-12">
            <label htmlFor="txtTitulo">Aplicação</label><span className={styles.required}> *</span>
            <table>
              <tr>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="aplicacao1" className="form-control" name="aplicacao1"
                      style={{ "height": "194px", "width": "300px" }}>
                      {this.state.itemsAplicacoes.map(function (item, key) {
                        return (
                          <option className="optAplicacao" value={item.ID}>{item.Title}</option>
                        );
                      })}
                    </select>
                  </div>
                </td>
                <td>
                  <div>
                    <input type="button" className="btn btn-light" id="addButtonAplicacao" onClick={this.addButtonAplicacao} value="Adicionar >" alt="Salvar" /></div><br />
                  <input type="button" className="btn btn-light" id="removeButtonAplicacao" onClick={this.removeButtonAplicacao} value="< Remover" alt="Salvar" />
                </td>
                <td>
                  <div className="col-sm-12">
                    <select multiple={true} id="aplicacao2" className="form-control" name="aplicacao2"
                      style={{ "height": "194px", "width": "300px" }}>
                    </select>
                  </div>
                </td>
              </tr>
            </table>
          </div>
        </div>


        <div className="form-group">
          <label htmlFor="txtDescricao">Descrição</label><span className={styles.required}> *</span>
          <textarea id="txtDescricao" className="form-control" rows={4}></textarea>
        </div>

        <div className="form-group">
          <label htmlFor="txtSolucao">Solução</label> <span className={styles.required}> *</span>
          <textarea id="txtSolucao" className="form-control" rows={4}></textarea>
        </div>

        <div className="form-group">
          <label htmlFor="txtObservacao">Observação</label>
          <textarea id="txtObservacao" className="form-control" rows={4}></textarea>
        </div>

        <div className="form-row">
          <div className="form-group col-md-6">

            <PeoplePicker
              context={this.props.context as any}
              titleText="Aprovador Engenharia"
              personSelectionLimit={3}
              groupName={"Aprovadores Engenharia"} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              required={true}
              disabled={false}
              onChange={this._getPeoplePickerItemsAprovadorEngenharia.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true} />
          </div>
          <div className="form-group col-md-6">
            <PeoplePicker
              context={this.props.context as any}
              titleText="Destinatários Adicionais"
              personSelectionLimit={3}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              required={false}
              disabled={false}
              onChange={this._getPeoplePickerItemsDestinatariosAdicionais.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true} />
          </div>
        </div>


        <div className="form-row">
          <div className="form-group col-md-6">

            <PeoplePicker
              context={this.props.context as any}
              titleText="Aprovador Geral"
              personSelectionLimit={3}
              groupName={"Aprovadores Gerais"} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              required={true}
              disabled={false}
              onChange={this._getPeoplePickerItemsAprovadorGeral.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true} />
          </div>
          <div className="form-group col-md-6">
          </div>
        </div>

        <div className="form-group">
          <div className="form-row">
            <div className="form-group col-md-6">
              <label htmlFor="ddlSegmento">Segmento</label><span className={styles.required}> *</span>
              <select id="ddlSegmento" className="form-control">
                <option value="0" selected>Selecione...</option>
                <option value="Normal">SST</option>
                <option value="Normal">OTHERS</option>
              </select>
            </div>
            <div className="form-group col-md-6">
            </div>
          </div>
        </div>

        <div className="form-group">
          <button id="btnIniciarAprovacao" className="btn btn-success" onClick={this.createNewItem}>Salvar</button>
        </div>
      </div>

        <div className="modal fade" id="modalSucesso" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                RC Cadastrada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="sucesso" onClick={this.fecharSucesso} className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>


        <div className="modal fade" id="modalValidacao" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                <span id="msgValidacao"></span>
              </div>
              <div className="modal-footer">
                <button className="btn btn-success" type="button" id="sucesso" onClick={this.fecharValidacao} >OK</button>
              </div>
            </div>
          </div>
        </div>

      </>


    );
  }


  //Eventos Produto

  addButtonProduto = () => {
    var $options = $('#produtos1 option:selected');
    $options.appendTo("#produtos2");
  }

  removeButtonProduto = () => {
    var $options = $('#produtos2 option:selected');
    $options.appendTo("#produtos1");
  }

  //Eventos Cliente

  addButtonCliente = () => {
    var $options = $('#cliente1 option:selected');
    $options.appendTo("#cliente2");
  }

  removeButtonCliente = () => {
    var $options = $('#cliente2 option:selected');
    $options.appendTo("#cliente1");
  }

  //Eventos Aplicação

  addButtonAplicacao = () => {
    var $options = $('#aplicacao1 option:selected');
    $options.appendTo("#aplicacao2");
  }

  removeButtonAplicacao = () => {
    var $options = $('#aplicacao2 option:selected');
    $options.appendTo("#aplicacao1");
  }


  _getPeoplePickerItemsAprovadorEngenharia(items: any[]) {
    console.log('Items:', items);
    this.setState({ addUsersAprovadorEngenharia: items as any });
    _addUsersAprovadorEngenharia = items;
  }

  _getPeoplePickerItemsAprovadorGeral(items: any[]) {
    console.log('Items:', items);
    this.setState({ addUsersAprovadorGeral: items as any });
    _addUsersAprovadorGeral = items;
  }

  _getPeoplePickerItemsDestinatariosAdicionais(items: any[]) {
    console.log('Items:', items);
    this.setState({ addUsersDestinatariosAdicionais: items as any });
    _addUsersDestinatariosAdicionais = items;
  }

  /*
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }
  */

  createNewItem() {

    try {

      $("#btnIniciarAprovacao").prop("disabled", true);

      var titulo = jQuery('#txtTitulo').val();
      var origembit = jQuery('#ddlOrigemBit option:selected').val();
      var descricao = jQuery('#txtDescricao').val();
      var solucao = jQuery('#txtSolucao').val();
      var observacao = jQuery('#txtObservacao').val();
      var segmento = jQuery('#ddlSegmento option:selected').val();
      var acoes = jQuery('#ddlAcoes option:selected').val();

      var vlrProduto = Array.prototype.slice.call(document.querySelectorAll('#produtos2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });

      var vlrCliente = Array.prototype.slice.call(document.querySelectorAll('#cliente2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });

      var vlrAplicacao = Array.prototype.slice.call(document.querySelectorAll('#aplicacao2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });


      var arrPeoplepickerAprovadorEngenharia = [];
      for (let i = 0; i < _addUsersAprovadorEngenharia.length; i++) {
        arrPeoplepickerAprovadorEngenharia.push(_addUsersAprovadorEngenharia[i]["id"]);
      }

      var arrPeoplepickerAprovadorGeral = [];
      for (let i = 0; i < _addUsersAprovadorGeral.length; i++) {
        arrPeoplepickerAprovadorGeral.push(_addUsersAprovadorGeral[i]["id"]);
      }

      var arrPeoplepickerDestinatariosAdicionais = [];
      for (let i = 0; i < _addUsersDestinatariosAdicionais.length; i++) {
        arrPeoplepickerDestinatariosAdicionais.push(_addUsersDestinatariosAdicionais[i]["id"]);
      }

      if (titulo == "") {
        alert("Forneça um título!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //jQuery("#msgValidacao").html("Forneça um título!");
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrProduto.length == 0) {
        alert("Selecione pelo menos um Produto!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Selecione pelo menos um Produto!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrCliente.length == 0) {
        alert("Selecione pelo menos um Cliente!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Selecione pelo menos um Cliente!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrAplicacao.length == 0) {
        alert("Selecione pelo menos uma Aplicação!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Selecione pelo menos uma Aplicação!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (descricao == "") {
        alert("Forneça uma descrição!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Forneça uma descrição!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (solucao == "") {
        alert("Forneça uma solução!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Forneça uma solução!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }


      if (_addUsersAprovadorEngenharia.length == 0) {
        alert("Forneça uma Aprovador Engenharia!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Forneça uma Aprovador Engenharia!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (_addUsersAprovadorGeral.length == 0) {
        alert("Forneça um Aprovador Geral!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Forneça uma Aprovador Suporte!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (segmento == 0) {
        alert("Selecione um Segmento!");
        $("#btnIniciarAprovacao").prop("disabled", false);
        //_msgValidacao = "Forneça uma Aprovador Suporte!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }


      var statusbit;
      var statusInterno;

      sp.web.lists
        .getByTitle("BIT")
        .items.add({
          Title: titulo,
          OrigemBIT: origembit,
          ProdutoId: { "results": vlrProduto },
          ClienteId: { 'results': vlrCliente },
          AplicacaoId: { 'results': vlrAplicacao },
          Descricao: descricao,
          Solucao: solucao,
          Observacao: observacao,
          Aprovador_x0020_EngenhariaId: { 'results': arrPeoplepickerAprovadorEngenharia },
          AprovadorGeralId: { 'results': arrPeoplepickerAprovadorGeral },
          Destinat_x00e1_rios_x0020_Padr_xId: { 'results': arrPeoplepickerDestinatariosAdicionais },
          Segmento: segmento,
          Acao: "Manter Em Elaboração",
          statusbit: "Em Elaboração",
          statusInterno: "Em Elaboração"
        })
        .then(response => {

          _idBit = response.data.ID;

          sp.web.lists.getByTitle("BIT").items.getById(response.data.ID).update({
            Title: response.data.ID + " - " + titulo,
            BITNumero: response.data.ID
          })
            .then(async response => {

              _web.lists.getByTitle("Anexos").rootFolder.folders.add(`${_idBit}`).then(data => {

                console.log("gravou2");
                jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });

              }).catch(err => {
                console.log("err",err);
              });

            });
        });
    } catch (ex) {
      console.log(ex);
      $("#btnIniciarAprovacao").prop("disabled", false);
      alert(ex);
    }

  }



  fecharSucesso() {

    $("#modalItens").modal('hide');
    window.location.href = `Editar-BIT.aspx?idBIT=` + _idBit;

  }


  fecharValidacao() {

    $("#modalValidacao").modal('hide');

  }


  mostraJustificativa() {

    var valor = $("#ddlAcoes option:checked").val();

    console.log(valor);

    if (valor == "Cancelar BIT") {
      $("#divJustificativaBit").show();
    } else {
      $("#divJustificativaBit").hide();
    }

  }



}



