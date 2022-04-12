import * as React from 'react';
import styles from './BitEditarItem.module.scss';
import { IBitEditarItemProps } from './IBitEditarItemProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';
import * as $ from "jquery";
import * as jQuery from "jquery";
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import "bootstrap";

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");

var testeGus;
var _addUsersAprovadorEngenharia = [];
var _addUsersAprovadorGeral = [];
var _addUsersDestinatariosAdicionais = [];
var _msgValidacao;
var _arrAprovadorEngenharia = [];
var _arrDestinatarios = [];
var _arrAprovadorGeral = [];
var _alterouPeoplePickerAprovadorEngenharia = false;
var _alterouPeoplePickerDestinatarios = false;
var _alterouPeoplePickerAprovadorGeral = false;
var _arrAprovadorEngenhariaID = [];
var _arrDestinatariosID = [];
var _arrAprovadorGeralID = [];
var _idBit;
var _userTitle;
var _size: number = 0;
var _web;
var _siteAntigo;
var _url;
var _bitNumero;
var _posBit = 0;

import { Web } from "sp-pnp-js";
import pnp from "sp-pnp-js";

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
  itemsAnexoItem: [
    {
      "ID": "",
      "ServerRelativeUrl": "",
      "FileName": "",
    }],
  showmessageBar: boolean; //to show/hide message bar on success
  message: string; // what message to be displayed in message bar
  itemID: number; // current item ID after create new item is clicked
  addUsersAprovadorEngenharia: [];
  addUsersAprovadorGeral: [];
  addUsersDestinatariosAdicionais: [];
  defaultmyusers: [];
  PeoplePickerDefaultItemsAprovadorEngenharia: string[];
  PeoplePickerDefaultItemsDestinatarios: string[];
  PeoplePickerDefaultItemsAprovadorGeral: string[];
}


export default class BitEditarItem extends React.Component<IBitEditarItemProps, IReactGetItemsState> {


  public constructor(props: IBitEditarItemProps, state: IReactGetItemsState) {
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
      itemsAnexoItem: [
        {
          "ID": "",
          "ServerRelativeUrl": "",
          "FileName": "",
        }],
      showmessageBar: false,
      message: "",
      itemID: 0,
      addUsersAprovadorEngenharia: [],
      addUsersAprovadorGeral: [],
      addUsersDestinatariosAdicionais: [],
      defaultmyusers: [],
      PeoplePickerDefaultItemsAprovadorEngenharia: [],
      PeoplePickerDefaultItemsDestinatarios: [],
      PeoplePickerDefaultItemsAprovadorGeral: []
    };
  }


  public async componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _url = this.props.siteurl;

    $("#divJustificativaBit").hide();
    $(".bitAntigo").hide();

    sp.web.currentUser.get().then(f => {
      _userTitle = f.Title;
      console.log("_userTitle2", _userTitle);
    })

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idBit = parseInt(queryParms.getValue("idBIT"));

    console.log("_idRCS", _idBit);

    document
      .getElementById("ddlAcoes")
      .addEventListener("change", (e: Event) => this.mostraJustificativa());


    document
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());

    var reactHandler = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Produtos')/items?$top=4999&$filter=Ativo eq 1&$orderby= ID desc`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandler.setState({
          items: resultData.d.results
        });
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
        reactHandler3.setState({
          itemsAplicacoes: resultData.d.results
        });
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
        reactHandler4.setState({
          itemsJustificativa: resultData.d.results
        });
        console.log("resultData.d.results4", resultData.d.results)
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('BIT')/items?$expand=Produto,Cliente,Aplicacao,Author,Aprovador_x0020_Engenharia,AprovadorGeral,Destinat_x00e1_rios_x0020_Padr_x&$select=ID,BITNumero,Title,OrigemBIT,Status,Produto/ID,Produto/Title,Cliente/ID,Cliente/Title,Aplicacao/ID,Aplicacao/Title,Segmento,Vers_x00e3_o_x0020_BIT,Author/Title,Created,Descricao,Solucao,Observacao,Aprovador_x0020_Engenharia/Title,AprovadorGeral/Title,Destinat_x00e1_rios_x0020_Padr_x/Title,Aprovador_x0020_Engenharia/ID,AprovadorGeral/ID,Destinat_x00e1_rios_x0020_Padr_x/ID,Segmento,Acao,SiteAntigo,txtAprovadorEngenharia,txtAprovadorGeral,txtDestinat_x00e1_riosAdicionais,BITNumero&$filter= ID eq ` + _idBit,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        var res = "";

        //var resultado = resultData.d.results;

        if (resultData.d.results.length > 0) {

          console.log();

          for (var i = 0; i < resultData.d.results.length; i++) {

            var txtProduto = "";
            var txtCliente = "";
            var txtAplicacao = "";

            for (var x = 0; x < resultData.d.results[i].Produto.results.length; x++) {

              txtProduto += `<option class="optProduto" value=${resultData.d.results[i].Produto.results[x].ID}>${resultData.d.results[i].Produto.results[x].Title}</option>`;
              $(`#produtos1 option[value='${resultData.d.results[i].Produto.results[x].ID}']`).remove();
            }

            for (var x = 0; x < resultData.d.results[i].Cliente.results.length; x++) {

              txtCliente += `<option class="optCliente" value=${resultData.d.results[i].Cliente.results[x].ID}>${resultData.d.results[i].Cliente.results[x].Title}</option>`;
              $(`#cliente1 option[value='${resultData.d.results[i].Cliente.results[x].ID}']`).remove();

            }

            for (var x = 0; x < resultData.d.results[i].Aplicacao.results.length; x++) {

              txtAplicacao += `<option class="optAplicacao" value=${resultData.d.results[i].Aplicacao.results[x].ID}>${resultData.d.results[i].Aplicacao.results[x].Title}</option>`;
              $(`#aplicacao1 option[value='${resultData.d.results[i].Aplicacao.results[x].ID}']`).remove();

            }

            var title = resultData.d.results[i].Title;
            var origemBit = resultData.d.results[i].OrigemBIT;
            var descricao = resultData.d.results[i].Descricao;
            var solucao = resultData.d.results[i].Solucao;
            var observacao = resultData.d.results[i].Observacao;
            var segmento = resultData.d.results[i].Segmento;
            var acoes = resultData.d.results[i].Acao;
            var bitNumero = resultData.d.results[i].BITNumero;

            _bitNumero = bitNumero;

            console.log("_bitNumero1", _bitNumero);

            $("#linkAnexo").html(`<a data-interception="off" target="_blank" href="${_url}/Anexos/${bitNumero}">Clique aqui</a> para fazer o upload do arquivo`);

            $("#txtTitulo").val(title);
            $("#ddlOrigemBit").val(origemBit);
            $("#produtos2").html(txtProduto);
            $("#cliente2").html(txtCliente);
            $("#aplicacao2").html(txtAplicacao);
            $("#txtDescricao").html(descricao);
            $("#txtSolucao").html(solucao);
            $("#txtObservacao").html(observacao);
            $("#ddlSegmento").val(segmento);
            $("#ddlAcoes").val(acoes);

            _siteAntigo = resultData.d.results[i].SiteAntigo;

            if (_siteAntigo) {

              $("#txtAprovadorEngenharia").html(resultData.d.results[i].txtAprovadorEngenharia);
              $("#txtDestinatarios").html(resultData.d.results[i].txtDestinat_x00e1_riosAdicionais);
              $("#txtAprovadorGeral").html(resultData.d.results[i].txtAprovadorGeral);
              $(".bitAntigo").show();

            }

            var id = resultData.d.results[i].ID;
            var bitNumero = resultData.d.results[i].BITNumero;
            var status = resultData.d.results[i].Status;

            var versaoBIT = resultData.d.results[i].Vers_x00e3_o_x0020_BIT;
            var author = resultData.d.results[i].Author.Title;
            var created = resultData.d.results[i].Created;


            if (resultData.d.results[i].Aprovador_x0020_Engenharia.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].Aprovador_x0020_Engenharia.results.length; x++) {

                _arrAprovadorEngenharia.push(resultData.d.results[i].Aprovador_x0020_Engenharia.results[x].Title);
                _arrAprovadorEngenhariaID.push(resultData.d.results[i].Aprovador_x0020_Engenharia.results[x].ID);


              }

            }


            if (resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results.length; x++) {

                _arrDestinatarios.push(resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results[x].Title);
                _arrDestinatariosID.push(resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results[x].ID);

              }

            }

            if (resultData.d.results[i].AprovadorGeral.hasOwnProperty('results')) {


              for (var x = 0; x < resultData.d.results[i].AprovadorGeral.results.length; x++) {

                _arrAprovadorGeral.push(resultData.d.results[i].AprovadorGeral.results[x].Title);
                _arrAprovadorGeralID.push(resultData.d.results[i].AprovadorGeral.results[x].ID);

              }

            }


          }

        }

      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


    this.GetDefaultUsers();


    //get anexos da biblioteca

    var montaAnexo = "";

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Editar-BIT.aspx", "");

    //var relative = "/sites/bit-hml";
    var idItem = 0;

    console.log("_bitNumero", _bitNumero);

    await _web.getFolderByServerRelativeUrl(`${strRelativeURL}/Anexos/${_bitNumero}`)
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
          montaAnexo = `<a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a> <a id="btnExcluirAnexo${idItem}" style="cursor:pointer" >Excluir</a><br/>`

          $("#conteudoAnexo").append(montaAnexo);

          document
            .getElementById(`btnExcluirAnexo${idItem}`)
            .addEventListener("click", (e: Event) => this.excluirAnexo(`${strRelativeURL}/Anexos/${_idBit}`, `${item.Name}`, `anexo${idItem}`, `btnExcluirAnexo${idItem}`));

        })

      });


    //fim anexos da biblioteca

    //get anexos de item (para os BITs do site antigo)

    var idItemAnexo = 0;
    var montaAnexoItem = "";

    var reactHandler5 = this;

    let item = await _web.lists.getByTitle("BIT").items.getById(_idBit);

    item.attachmentFiles.get().then(v => {

      reactHandler5.setState({
        itemsAnexoItem: v
      });

    });


    //fim anexos do item




  }



  private newMethod() {
    return this;
  }

  private GetDefaultUsers() {

    setTimeout(() => {

      console.log("_arrAprovadorEngenharia1", _arrAprovadorEngenharia);
      console.log("_arrDestinatarios1", _arrDestinatarios);
      console.log("_arrAprovadorGeral1", _arrAprovadorGeral);

      console.log("_arrAprovadorEngenhariaID", _arrAprovadorEngenhariaID);
      console.log("_arrDestinatariosID", _arrDestinatariosID);
      console.log("_arrAprovadorGeralID", _arrAprovadorGeralID);

      this.setState({
        PeoplePickerDefaultItemsAprovadorEngenharia: _arrAprovadorEngenharia,
        PeoplePickerDefaultItemsDestinatarios: _arrDestinatarios,
        PeoplePickerDefaultItemsAprovadorGeral: _arrAprovadorGeral
      });

    }, 2000);


  }


  public render(): React.ReactElement<IBitEditarItemProps> {

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
                <option value="Engenharia AT">Engenharia AT</option>
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

        <div className="form-group bitAntigo">
          <b>Aprovadores - BITs antigos</b>
        </div>

        <div className="form-row bitAntigo">
          <div className="form-group col-md-6">
            <label htmlFor="txtAprovadorEngenharia">Aprovador Engenharia</label>
            <br></br><span className="text-info" id="txtAprovadorEngenharia"></span>
          </div>
          <div className="form-group col-md-6">
            <label htmlFor="txtDestinatarios">Destinatários Adicionais</label>
            <br></br><span className="text-info" id="txtDestinatarios"></span>
          </div>
        </div>

        <div className="form-group bitAntigo">
          <label htmlFor="txtAprovadorGeral">Aprovador Geral</label>
          <br></br><span className="text-info" id="txtAprovadorGeral"></span>
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
              defaultSelectedUsers={this.state.PeoplePickerDefaultItemsAprovadorEngenharia}
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
              defaultSelectedUsers={this.state.PeoplePickerDefaultItemsDestinatarios}
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
              defaultSelectedUsers={this.state.PeoplePickerDefaultItemsAprovadorGeral}
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
              <label htmlFor="ddlAcoes">Ações</label><span className={styles.required}> *</span>
              <select id="ddlAcoes" className="form-control">
                <option value="-">Selecione</option>
                <option value="Manter Em Elaboração">Manter Em Elaboração</option>
                <option value="Submeter a Aprovação" >Submeter a Aprovação</option>
                <option value="Cancelar BIT" >Cancelar BIT</option>
              </select>
            </div>
          </div>
        </div>

        <div className="form-row" id="divJustificativaBit">
          <div className="form-group col-md-6">
            <label htmlFor="ddlAcoes">Justificativa - Cancelar Bit</label><span className={styles.required}> *</span>
            <select id="ddlJustificarBit" className="form-control">
              <option value="0" selected>Selecione...</option>
              {this.state.itemsJustificativa.map(function (item, key) {
                return (
                  <option className="optAplicacoes" value={item.ID}>{item.Title}</option>
                );
              })}
            </select>
          </div>
          <div className="form-group col-md-6">
          </div>
        </div>

        <div className="form-group">
          <div className="form-row">
            <div className="form-group col-md-10">
              <label>Anexos</label>
              <br></br>

              <span id="linkAnexo"></span>

              <br></br><br></br>
              <div id="conteudoAnexoNaoEncontrado"><p>Nenhum anexo encontrado</p></div>
              <div id="conteudoAnexo"></div><br></br>
              <div id="conteudoAnexo2">

                {this.state.itemsAnexoItem.map((item, key) => {

                  console.log("item anexos",item)

                  _posBit++;
                  var txtAnexoItem = "anexoItem" + _posBit;
                  var btnExcluirAnexoitem = "btnExcluirAnexoitem" + _posBit;
                  $("#conteudoAnexoNaoEncontrado").hide();

                  return (
                    <><a id={txtAnexoItem} data-interception='off' target='_blank' href={item.ServerRelativeUrl}>{item.FileName}</a>
                      &nbsp;<a onClick={() => this.excluirAnexoItem(`${item.ServerRelativeUrl}`, `${item.FileName}`, `${txtAnexoItem}`, `${btnExcluirAnexoitem}`)}  id={btnExcluirAnexoitem} style={{ "cursor": "pointer" }}>Excluir</a><br /></>
                  );


                })}


              </div>

            </div>

            <div className="form-group col-md-2">
            </div>
          </div>
        </div>

        <div className="form-group">
          <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
          <button id="btnIniciarAprovacao" className="btn btn-success" onClick={this.Update}>Salvar</button>
        </div>
      </div>

        <div className="modal fade" id="modalSucesso" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                RC atualizada com sucesso!
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

  async excluirAnexoItem(ServerRelativeUr, name, elemento, elemento2)
{         
  

  if (confirm("Deseja realmente excluir o arquivo " + name + "?") == true) {

    var relativeURL = window.location.pathname;
    var strRelativeURL = relativeURL.replace("SitePages/Editar-BIT.aspx", "");

    console.log("(`${strRelativeURL}/Lists/bit/Attachments/${_bitNumero}`)",(`${strRelativeURL}/Lists/bit/Attachments/${_bitNumero}`))

    await _web.getFolderByServerRelativeUrl(`${strRelativeURL}/Lists/bit/Attachments/${_idBit}`).files.getByName(name).delete()
      .then(async response => {
        jQuery(`#${elemento}`).hide();
        jQuery(`#${elemento2}`).hide();
        alert("Arquivo excluido com sucesso.");
      }).catch(console.error());

  } else {
    return false;
  }
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
    _alterouPeoplePickerAprovadorEngenharia = true;
  }

  _getPeoplePickerItemsAprovadorGeral(items2: any[]) {
    console.log('Items2:', items2);
    this.setState({ addUsersAprovadorGeral: items2 as any });
    _addUsersAprovadorGeral = items2;
    _alterouPeoplePickerAprovadorGeral = true;
  }

  _getPeoplePickerItemsDestinatariosAdicionais(items3: any[]) {
    console.log('Items3:', items3);
    this.setState({ addUsersDestinatariosAdicionais: items3 as any });
    _addUsersDestinatariosAdicionais = items3;
    _alterouPeoplePickerDestinatarios = true;
  }

  /*
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }
*/

  private Update() {

    try {

      console.log("entrou update");

      $('.optProduto').prop('selected', true);
      $('.optCliente').prop('selected', true);
      $('.optAplicacao').prop('selected', true);

      var titulo = jQuery('#txtTitulo').val();
      var origembit = jQuery('#ddlOrigemBit option:selected').val();
      var descricao = jQuery('#txtDescricao').val();
      var solucao = jQuery('#txtSolucao').val();
      var observacao = jQuery('#txtObservacao').val();
      var segmento = jQuery('#ddlSegmento option:selected').val();
      var acoes = jQuery('#ddlAcoes option:selected').val();
      var justificarBit = jQuery('#ddlJustificarBit option:selected').val();
      

      var vlrProduto = Array.prototype.slice.call(document.querySelectorAll('#produtos2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });

      var vlrCliente = Array.prototype.slice.call(document.querySelectorAll('#cliente2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });

      var vlrAplicacao = Array.prototype.slice.call(document.querySelectorAll('#aplicacao2 option:checked'), 0).map(function (v, i, a) {
        return v.value;
      });


      console.log("_addUsersAprovadorEngenharia0", _addUsersAprovadorEngenharia);
      var arrPeoplepickerAprovadorEngenharia = [];
      for (let i = 0; i < _addUsersAprovadorEngenharia.length; i++) {
        arrPeoplepickerAprovadorEngenharia.push(_addUsersAprovadorEngenharia[i]["id"]);
      }

      console.log("_addUsersAprovadorGeral0", _addUsersAprovadorGeral);
      var arrPeoplepickerAprovadorGeral = [];
      for (let i = 0; i < _addUsersAprovadorGeral.length; i++) {
        arrPeoplepickerAprovadorGeral.push(_addUsersAprovadorGeral[i]["id"]);
      }

      console.log("_addUsersDestinatariosAdicionais0", _addUsersDestinatariosAdicionais);
      var arrPeoplepickerDestinatariosAdicionais = [];
      for (let i = 0; i < _addUsersDestinatariosAdicionais.length; i++) {
        arrPeoplepickerDestinatariosAdicionais.push(_addUsersDestinatariosAdicionais[i]["id"]);
      }

      if (titulo == "") {
        alert("Forneça um título!");
        //jQuery("#msgValidacao").html("Forneça um título!");
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrProduto.length == 0) {
        alert("Selecione pelo menos um Produto!");
        //_msgValidacao = "Selecione pelo menos um Produto!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrCliente.length == 0) {
        alert("Selecione pelo menos um Cliente!");
        //_msgValidacao = "Selecione pelo menos um Cliente!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (vlrAplicacao.length == 0) {
        alert("Selecione pelo menos uma Aplicação!");
        //_msgValidacao = "Selecione pelo menos uma Aplicação!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (descricao == "") {
        alert("Forneça uma descrição!");
        //_msgValidacao = "Forneça uma descrição!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (solucao == "") {
        alert("Forneça uma solução!");
        //_msgValidacao = "Forneça uma solução!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }


      console.log("arrPeoplepickerAprovadorEngenharia", arrPeoplepickerAprovadorEngenharia);
      console.log("arrPeoplepickerAprovadorGeral", arrPeoplepickerAprovadorGeral);
      console.log("arrPeoplepickerDestinatariosAdicionais", arrPeoplepickerDestinatariosAdicionais);

      var resAprovadorGeral;
      var resDestinatarios;
      var resAprovadorEngenharia;

      if (_alterouPeoplePickerAprovadorEngenharia) {

        resAprovadorEngenharia = arrPeoplepickerAprovadorEngenharia;

      }
      else {

        resAprovadorEngenharia = _arrAprovadorEngenhariaID;
      }

      //

      if (_alterouPeoplePickerDestinatarios) {

        resDestinatarios = arrPeoplepickerDestinatariosAdicionais;

      }
      else {

        resDestinatarios = _arrDestinatariosID;
      }


      if (_alterouPeoplePickerAprovadorGeral) {

        resAprovadorGeral = arrPeoplepickerAprovadorGeral;

      }
      else {

        resAprovadorGeral = _arrAprovadorGeralID;
      }

      if (resAprovadorEngenharia.length == 0) {
        alert("Forneça uma Aprovador Engenharia!");
        //_msgValidacao = "Forneça uma Aprovador Engenharia!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (resAprovadorGeral.length == 0) {
        alert("Forneça um Aprovador Geral!");
        //_msgValidacao = "Forneça uma Aprovador Suporte!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (segmento == 0) {
        alert("Selecione um Segmento!");
        //_msgValidacao = "Forneça uma Aprovador Suporte!";
        //jQuery("#modalValidacao").modal({backdrop: 'static', keyboard: false});
        return false;
      }

      if (acoes == "-") {

        alert("Escolha uma ação!");
        return false;

      }

      var vlrJustificarBit = "";

      if (acoes == "Cancelar BIT") {

        if (justificarBit == "0") {
          alert("Forneça uma justificativa para cancelar o Bit!");
          return false;
        }

        vlrJustificarBit = jQuery('#ddlJustificarBit option:selected').text();

      }

      var statusbit;
      var statusInterno;

      if (acoes == "Manter Em Elaboração") {
        statusbit = "Em Elaboração";
        statusInterno = "Em Elaboração";
      }
      else if (acoes == "Submeter a Aprovação") {
        statusbit = "Aguardando Aprovação";
        statusInterno = "Aguardando Aprovação Engenharia";
      }
      else if (acoes == "Cancelar BIT") {
        statusbit = "Cancelado";
        statusInterno = "Cancelado";
      }

      sp.web.lists
        .getByTitle("BIT")
        .items.getById(_idBit).update({
          Title: titulo,
          OrigemBIT: origembit,
          ProdutoId: { "results": vlrProduto },
          ClienteId: { 'results': vlrCliente },
          AplicacaoId: { 'results': vlrAplicacao },
          Descricao: descricao,
          Solucao: solucao,
          Observacao: observacao,
          Aprovador_x0020_EngenhariaId: { 'results': resAprovadorEngenharia },
          AprovadorGeralId: { 'results': resAprovadorGeral },
          Destinat_x00e1_rios_x0020_Padr_xId: { 'results': resDestinatarios },
          Segmento: segmento,
          Acao: "-",
          statusbit: statusbit,
          statusInterno: statusInterno,
          SiteAntigo: false,
          JustificativaCancelarBit: vlrJustificarBit

        })
        .then(async response => {

          if (acoes == "Submeter a Aprovação") {

            var today = new Date();
            var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
            var date = today.getDate() + '/' + (today.getMonth() + 1) + '/' + today.getFullYear();

            sp.web.lists
              .getByTitle("Historico")
              .items.add({
                Title: _userTitle + " submeteu o BIT para aprovação em " + date + " as " + time,
                BITId: _idBit,
                TemplateEmail: "Aguardando Aprovação Engenharia"
              })
              .then(async response => {

                console.log("gravou");
                jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });

              });

          }
          else if (acoes == "Cancelar BIT") {

            var today = new Date();
            var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
            var date = today.getDate() + '/' + (today.getMonth() + 1) + '/' + today.getFullYear();

            sp.web.lists
              .getByTitle("Historico")
              .items.add({
                Title: _userTitle + " cancelou o BIT em " + date + " as " + time,
                BITId: _idBit,
                TemplateEmail: "Cancelado"
              })
              .then(async response => {

                console.log("gravou");
                jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });


              });

          }


          else {

            console.log("gravou");
            jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });

          }


        });



    } catch (ex) {
      console.log(ex);
      alert(ex);
    }



  }


  fecharSucesso() {

    $("#modalItens").modal('hide');
    window.location.href = `BIT.aspx`;

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

  voltar() {

    history.back();

  }

  public async excluirAnexoItem2(name, elemento, elemento2) {

    if (confirm("Deseja realmente excluir o arquivo " + name + "?") == true) {

      //let item = _web.lists.getByTitle("BIT").items.getById(_idBit);
      //item.attachmentFiles.getByName("file2.txt").delete().then(v => {
      // jQuery(`#${elemento}`).hide();
      // jQuery(`#${elemento2}`).hide();
      //  alert("Arquivo excluido com sucesso.");
      //  })

    } else {
      return false;
    }


  }

  async excluirAnexo(ServerRelativeUr, name, elemento, elemento2) {

    if (confirm("Deseja realmente excluir o arquivo " + name + "?") == true) {

      console.log("ServerRelativeUr", ServerRelativeUr);
      console.log("name", name);
      await _web.getFolderByServerRelativeUrl(ServerRelativeUr).files.getByName(name).delete()
        .then(async response => {
          jQuery(`#${elemento}`).hide();
          jQuery(`#${elemento2}`).hide();
          alert("Arquivo excluido com sucesso.");
        }).catch(console.error());

    } else {
      return false;
    }

  }


}
