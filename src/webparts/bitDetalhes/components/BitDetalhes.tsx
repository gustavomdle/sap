import * as React from 'react';
import styles from './BitDetalhes.module.scss';
import { IBitDetalhesProps } from './IBitDetalhesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';
import * as $ from "jquery";
import * as jQuery from "jquery";
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import "bootstrap";
import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import { ICamlQuery } from '@pnp/sp/lists';
import { Web } from "sp-pnp-js";

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");

var _idBit;
var _userTitle;
var _arrAprovadorEngenharia = [];
var _arrAprovadorGeral = [];
var _aprovadorEngenharia = false;
var _aprovadorGeral = false;
var _web;
var _statusInterno;
var _versao = 0;
var _bitNumero;
var _url;



export default class BitDetalhes extends React.Component<IBitDetalhesProps, {}> {


  public async componentDidMount() {

    var _testeGus;

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    // let groups = await _web.currentUser.groups();

    // console.log("groups",groups);

    jQuery("#btnReprovar").hide();
    jQuery("#btnEditar").hide();
    jQuery("#btnAprovar").hide();
    jQuery("#conteudoJustificativa").hide();

    $("#trJustificativaCancelamento").hide();

    document
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());

    document
      .getElementById("btnEditar")
      .addEventListener("click", (e: Event) => this.enviarParaEdicao());


    document
      .getElementById("btnReprovar")
      .addEventListener("click", (e: Event) => this.reprovar());


    document
      .getElementById("btnAprovar")
      .addEventListener("click", (e: Event) => this.aprovar());


    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idBit = parseInt(queryParms.getValue("idBIT"));


    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('BIT')/items?$expand=Produto,Cliente,Aplicacao,Author,Aprovador_x0020_Engenharia,AprovadorGeral,Destinat_x00e1_rios_x0020_Padr_x&$select=ID,Title,OrigemBIT,Descricao,Solucao,Observacao,Aprovador_x0020_Engenharia/Title,AprovadorGeral/Title,Destinat_x00e1_rios_x0020_Padr_x/Title,BITNumero,Status,Produto/Title,Cliente/Title,Aplicacao/Title,Segmento,Acao,Vers_x00e3_o_x0020_BIT,Author/Title,Created,statusbit,Aprova_x00e7__x00e3_o_x0020_Enge,Aprova_x00e7__x00e3_o_x0020_Gera,statusInterno,SiteAntigo,txtAprovadorEngenharia,txtAprovadorGeral,txtDestinat_x00e1_riosAdicionais,JustificativaCancelarBit,BITNumero&$filter= ID eq ` + _idBit,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        //var resultado = resultData.d.results;

        if (resultData.d.results.length > 0) {

          console.log();

          for (var i = 0; i < resultData.d.results.length; i++) {

            var txtProduto = "";
            var txtCliente = "";
            var txtAplicacao = "";
            var txtAprovadorEngenharia = "";
            var txtDestinatarios = "";
            var txtAprovadorGeral = "";

            var siteAntigo = resultData.d.results[i].SiteAntigo;


            for (var x = 0; x < resultData.d.results[i].Produto.results.length; x++) {

              if (x == resultData.d.results[i].Produto.results.length - 1) {

                txtProduto += resultData.d.results[i].Produto.results[x].Title;

              } else {

                txtProduto += resultData.d.results[i].Produto.results[x].Title + ", ";
              }

            }


            for (var x = 0; x < resultData.d.results[i].Cliente.results.length; x++) {

              if (x == resultData.d.results[i].Cliente.results.length - 1) {

                txtCliente += resultData.d.results[i].Cliente.results[x].Title;

              } else {

                txtCliente += resultData.d.results[i].Cliente.results[x].Title + ", ";
              }
            }


            for (var x = 0; x < resultData.d.results[i].Aplicacao.results.length; x++) {

              if (x == resultData.d.results[i].Aplicacao.results.length - 1) {

                txtAplicacao += resultData.d.results[i].Aplicacao.results[x].Title;
              } else {

                txtAplicacao += resultData.d.results[i].Aplicacao.results[x].Title + ", ";
              }
            }


            if (resultData.d.results[i].Aprovador_x0020_Engenharia.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].Aprovador_x0020_Engenharia.results.length; x++) {

                _arrAprovadorEngenharia.push(resultData.d.results[i].Aprovador_x0020_Engenharia.results[x].Title);

                if (x == resultData.d.results[i].Aprovador_x0020_Engenharia.results.length - 1) {

                  txtAprovadorEngenharia += resultData.d.results[i].Aprovador_x0020_Engenharia.results[x].Title;

                } else {

                  txtAprovadorEngenharia += resultData.d.results[i].Aprovador_x0020_Engenharia.results[x].Title + ", ";
                }

              }

            }

            if (resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results.length; x++) {

                if (x == resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results.length - 1) {

                  txtDestinatarios += resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results[x].Title;

                } else {

                  txtDestinatarios += resultData.d.results[i].Destinat_x00e1_rios_x0020_Padr_x.results[x].Title + ", ";
                }

              }

            }


            if (resultData.d.results[i].AprovadorGeral.hasOwnProperty('results')) {


              for (var x = 0; x < resultData.d.results[i].AprovadorGeral.results.length; x++) {

                _arrAprovadorGeral.push(resultData.d.results[i].AprovadorGeral.results[x].Title);

                if (x == resultData.d.results[i].AprovadorGeral.results.length - 1) {

                  txtAprovadorGeral += resultData.d.results[i].AprovadorGeral.results[x].Title;

                } else {

                  txtAprovadorGeral += resultData.d.results[i].AprovadorGeral.results[x].Title + ", ";
                }

              }

            }



            console.log("txtAprovadorGeral", txtAprovadorGeral);

            //var bitNumero = resultData.d.results[i].BITNumero;
            var id = resultData.d.results[i].ID;
            var origemBIT = resultData.d.results[i].OrigemBIT;
            var title = resultData.d.results[i].Title;
            var descricao = resultData.d.results[i].Descricao;
            var solucao = resultData.d.results[i].Solucao;
            var observacao = resultData.d.results[i].Observacao;
            var segmento = resultData.d.results[i].Segmento;
            var acoes = resultData.d.results[i].Acao;
            var bitNumero = resultData.d.results[i].BITNumero;

            _bitNumero = bitNumero;

            console.log("_bitNumero1", _bitNumero);

            var status = resultData.d.results[i].statusbit;
            var versaoBIT = resultData.d.results[i].Vers_x00e3_o_x0020_BIT;
            var justificativaCancelarBit = resultData.d.results[i].JustificativaCancelarBit;
            _versao = versaoBIT;
            var aprovacaoEngenharia = resultData.d.results[i].Aprova_x00e7__x00e3_o_x0020_Enge;
            var aprovacaoGeral = resultData.d.results[i].Aprova_x00e7__x00e3_o_x0020_Gera;
            var statusInterno = resultData.d.results[i].statusInterno;
            _statusInterno = statusInterno;


            if (_arrAprovadorEngenharia.indexOf(_userTitle) !== -1) {

              _aprovadorEngenharia = true;
              console.log("_aprovadorEngenharia", _aprovadorEngenharia)

            }

            if (_arrAprovadorGeral.indexOf(_userTitle) !== -1) {

              _aprovadorGeral = true;
              console.log("_aprovadorGeral", _aprovadorGeral)

            }

            if (title == null) title = "";
            if (title == null) title = "";

            $("#txtTitulo").html(title);
            $("#txtOrigemBit").html(origemBIT);
            $("#txtProduto").html(txtProduto);
            $("#txtCliente").html(txtCliente);
            $("#txtAplicacao").html(txtAplicacao);
            $("#txtDescricao").html(descricao);
            $("#txtSolucao").html(solucao);
            $("#txtObservacao").html(observacao);
            $("#txtJustificativaCancelamento").html(justificativaCancelarBit);

            console.log("siteAntigo", siteAntigo);

            if (siteAntigo) {

              $("#txtAprovadorEngenharia").html(resultData.d.results[i].txtAprovadorEngenharia);
              $("#txtDestinatarios").html(resultData.d.results[i].txtDestinat_x00e1_riosAdicionais);
              $("#txtAprovadorGeral").html(resultData.d.results[i].txtAprovadorGeral);

            } else {

              $("#txtAprovadorEngenharia").html(txtAprovadorEngenharia);
              $("#txtDestinatarios").html(txtDestinatarios);
              $("#txtAprovadorGeral").html(txtAprovadorGeral);

            }

            $("#txtAcoes").html(acoes);
            $("#txtSegmento").html(segmento);
            $("#txtStatus").html(status);
            $("#txtVersao").html(versaoBIT);

            _web.currentUser.get().then(f => {

              _userTitle = f.Title;

              if (_arrAprovadorEngenharia.indexOf(_userTitle) !== -1) {

                _aprovadorEngenharia = true;

              }

              if (_arrAprovadorGeral.indexOf(_userTitle) !== -1) {

                _aprovadorGeral = true;

              }

              console.log("_userTitle", _userTitle);

              console.log("_arrAprovadorEngenharia", _arrAprovadorEngenharia);
              console.log("_aprovadorEngenharia", _aprovadorEngenharia);
              console.log("_arrAprovadorGeral", _arrAprovadorGeral);
              console.log("_aprovadorGeral", _aprovadorGeral);


              if (status == "Cancelado") {
                $("#trJustificativaCancelamento").show();
              }

              if (status == "Aguardando Aprovação") {

                if (statusInterno == "Aguardando Aprovação Engenharia" && (_aprovadorEngenharia)) {

                  jQuery("#btnReprovar").show();
                  jQuery("#btnAprovar").show();
                  jQuery("#conteudoJustificativa").show();

                }

                if (statusInterno == "Aguardando Aprovação Final" && (_aprovadorGeral)) {

                  jQuery("#btnReprovar").show();
                  jQuery("#btnAprovar").show();
                  jQuery("#conteudoJustificativa").show();

                }

              }

            })

          }

        }


      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    var montaItem = "";

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Historico')/items?$select=ID,Title,Justificativa&$filter= BIT eq ` + _idBit,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        //var resultado = resultData.d.results;

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var id = resultData.d.results[i].ID;
            var justificativa = resultData.d.results[i].Justificativa;
            if (justificativa == null) justificativa = "";

            console.log("id", id);

            montaItem += `<tr>
            <td style="max-width: 500px; width: 500px">${resultData.d.results[i].Title}</td>
            <td style="max-width: 300px; width: 300px">${justificativa}</td>
            </tr>`;

          }

          $("#conteudoHistorico").html(montaItem);

        }

      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

    //get anexos de item (para os BITs do site antigo)

    var montaAnexoItem = "";
    var url = `${this.props.siteurl}/_api/web/lists/getByTitle('BIT')/items('${_idBit}')/AttachmentFiles`;
    var _url = this.props.siteurl;
    console.log("url", url);

    $.ajax
      ({
        url: url,
        method: "GET",
        headers:
        {
          // Accept header: Specifies the format for response data from the server.
          "Accept": "application/json;odata=verbose"
        },
        success: function (data, status, xhr) {
          var dataresults = data.d.results;

          // _testeGus = data.d.results;

          console.log("dataresults", dataresults);

          for (var i = 0; i < dataresults.length; i++) {

            $("#conteudoAnexoNaoEncontrado").hide();

            montaAnexoItem += `<a title="" href="${_url}/Lists/bit/Attachments/${_idBit}/${dataresults[i]["FileName"]}">${dataresults[i]["FileName"]}</a><br/>`

          }
          $("#conteudoAnexo2").html(montaAnexoItem);
        },
        error: function (xhr, status, error) {
          console.log("Falha anexo");
        }
      }).catch((error: any) => {
        console.log("Erro Anexo do item: ", error);
      });

    //fim anexos do item

    //get anexos da biblioteca

    var montaAnexo = "";

    var relativeURL = window.location.pathname;

    var strRelativeURL = relativeURL.replace("SitePages/Detalhes-BIT.aspx", "");

    //var relative = "/sites/bit-hml";
    var idItem = 0;

    console.log("_bitNumero", _bitNumero);

    console.log("caminho",`${strRelativeURL}/Anexos/${_bitNumero}`);


    _web.getFolderByServerRelativeUrl(`${strRelativeURL}/Anexos/${_bitNumero}`)
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
          montaAnexo = `<a id="anexo${idItem}" data-interception="off" target="_blank" title="" href="${item.ServerRelativeUrl}">${item.Name}</a> <br/>`

          $("#conteudoAnexo").append(montaAnexo);

        })

      }).catch((error: any) => {
        console.log("Erro Anexo da biblioteca: ", error);
      });


    //fim anexos da biblioteca




  }


  public render(): React.ReactElement<IBitDetalhesProps> {
    return (


      <><div>

        <div className="container-fluid border" style={{ "width": "840px" }}>

          <div className="form-group">
            <div className="form-row">
              <div className="form-group col-md-6">
                <label htmlFor="txtTitulo">Título</label>
                <br></br><span className="text-info" id="txtTitulo"></span>
              </div>
              <div className="form-group col-md-4">
                <label htmlFor="txtStatus">Status</label>
                <br></br><span className="text-info" id="txtStatus"></span>
              </div>
              <div className="form-group col-md-2">
                <label htmlFor="txtStatus">Versão</label>
                <br></br><span className="text-info" id="txtVersao"></span>
              </div>
            </div>
          </div>

          <div id='trJustificativaCancelamento' className="form-group">
            <label htmlFor="txtJustificativaCancelamento">Justificativa de Cancelamento</label>
            <br></br><span className="text-info" id="txtJustificativaCancelamento"></span>
          </div>


          <div className="form-group">
            <label htmlFor="txtOrigemBit">Origem BIT</label>
            <br></br><span className="text-info" id="txtOrigemBit"></span>
          </div>


          <div className="form-group">
            <label htmlFor="txtProduto">Produto</label>
            <br></br><span className="text-info" id="txtProduto"></span>
          </div>

          <div className="form-group">
            <label htmlFor="txtCliente">Cliente</label>
            <br></br><span className="text-info" id="txtCliente"></span>
          </div>

          <div className="form-group">
            <label htmlFor="txtAplicacao">Aplicação</label>
            <br></br><span className="text-info" id="txtAplicacao"></span>
          </div>


          <div className="form-group">
            <label htmlFor="txtDescricao">Descrição</label>
            <br></br><span className="text-info" id="txtDescricao"></span>
          </div>

          <div className="form-group">
            <label htmlFor="txtSolucao">Solução</label>
            <br></br><span className="text-info" id="txtSolucao"></span>
          </div>

          <div className="form-group">
            <label htmlFor="txtObservacao">Observação</label>
            <br></br><span className="text-info" id="txtObservacao"></span>
          </div>

          <div className="form-row">
            <div className="form-group col-md-6">
              <label htmlFor="txtAprovadorEngenharia">Aprovador Engenharia</label>
              <br></br><span className="text-info" id="txtAprovadorEngenharia"></span>
            </div>
            <div className="form-group col-md-6">
              <label htmlFor="txtDestinatarios">Destinatários Adicionais</label>
              <br></br><span className="text-info" id="txtDestinatarios"></span>
            </div>
          </div>

          <div className="form-group">
            <label htmlFor="txtAprovadorGeral">Aprovador Geral</label>
            <br></br><span className="text-info" id="txtAprovadorGeral"></span>
          </div>

          <div className="form-group">
            <div className="form-row">
              <div className="form-group col-md-6">
                <label htmlFor="txtSegmento">Segmento</label>
                <br></br><span className="text-info" id="txtSegmento"></span>
              </div>
              <div className="form-group col-md-6">

              </div>
            </div>
          </div>

          <div className="form-group">
            <div className="form-row">
              <div className="form-group col-md-10">
                <label>Anexos</label>

                <div id="conteudoAnexoNaoEncontrado"><p>Nenhum anexo encontrado</p></div>
                <div id="conteudoAnexo"></div><br></br>
                <div id="conteudoAnexo2"></div>

              </div>

              <div className="form-group col-md-2">
              </div>
            </div>
          </div>

        </div>

      </div>

        <br></br><br></br>

        <div className="container-fluid border">

          <h4>HISTÓRICO</h4>

          <div className="table-responsive">
            <table className="table table-hover" id="tbItens">
              <thead>
                <tr>
                  <th scope="col">Ação</th>
                  <th scope="col">Justificativa</th>
                </tr>
              </thead>
              <tbody id="conteudoHistorico">
              </tbody>
            </table>
          </div>

        </div>

        <br></br><br></br>

        <div className="table-responsive" id="conteudoJustificativa">
          <label htmlFor="txtJustificativa">Justificativa</label>
          <textarea id="txtJustificativa" className="form-control" rows={4}></textarea>
        </div>


        <br></br><br></br>

        <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
        <button style={{ "margin": "2px" }} type="submit" id="btnEditar" className="btn btn-primary">Editar</button>
        <button style={{ "margin": "2px" }} type="submit" id="btnReprovar" className="btn btn-danger">Reprovar</button>
        <button style={{ "margin": "2px" }} type="submit" id="btnAprovar" className="btn btn-success">Aprovar</button>


        <div className="modal fade" id="modalAprovado" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                BIT aprovado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="sucesso" onClick={this.fecharSucesso} className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalReprovado" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                BIT reprovado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="sucesso" onClick={this.fecharSucesso} className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

      </>


    );
  }


  enviarParaEdicao() {

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _idBit = parseInt(queryParms.getValue("idBIT"));

    window.location.href = `Editar-BIT.aspx?idBIT=${_idBit}`;

  }


  async aprovar() {

    $("#btnReprovar").prop("disabled", true);
    $("#btnEditar").prop("disabled", true);
    $("#btnAprovar").prop("disabled", true);
    $("#btnVoltar").prop("disabled", true);

    var r = confirm("Deseja realmente aprovar esse BIT?");

    if (r == true) {

      try {

        var textoJustificativaHistorico = $("#txtJustificativa").val();

        var today = new Date();
        var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
        var date = today.getDate() + '/' + (today.getMonth() + 1) + '/' + today.getFullYear();

        if (_statusInterno == "Aguardando Aprovação Engenharia") {

          await _web.lists.getByTitle("BIT").items.getById(_idBit).update({
            statusInterno: "Aguardando Aprovação Final",
          }).then(async response => {

            _web.lists
              .getByTitle("Historico")
              .items.add({
                Title: _userTitle + " aprovou a tarefa Aprovação Engenharia  " + date + " as " + time,
                BITId: _idBit,
                Justificativa: textoJustificativaHistorico,
                TemplateEmail: "Aguardando Aprovação Final"

              })
              .then(response => {

                console.log("aprovou");
                jQuery("#modalAprovado").modal({ backdrop: 'static', keyboard: false });


              });

          }).catch((error: any) => {
            console.log("Error: ", error);
            //this.criarLog(error.responseJSON.detailedMessage);
          });


        }


        else if (_statusInterno == "Aguardando Aprovação Final") {

          var versaoNova = _versao + 1;

          await _web.lists.getByTitle("BIT").items.getById(_idBit).update({
            statusInterno: "Aprovado",
            statusbit: "Aprovado",
            Vers_x00e3_o_x0020_BIT: versaoNova,
            Acao: "-"

          }).then(async response => {

            _web.lists
              .getByTitle("Historico")
              .items.add({
                Title: _userTitle + " aprovou a tarefa Aprovação Final  " + date + " as " + time,
                BITId: _idBit,
                Justificativa: textoJustificativaHistorico,
                TemplateEmail: "Aprovado"

              })
              .then(response => {

                console.log("aprovou");
                jQuery("#modalAprovado").modal({ backdrop: 'static', keyboard: false });


              });

          }).catch((error: any) => {
            console.log("Error: ", error);
            //this.criarLog(error.responseJSON.detailedMessage);
          });


        }


      } catch (ex) {
        console.log(ex);;
        // this.criarLog(ex);
      }


    } else {
      $("#btnReprovar").prop("disabled", false);
      $("#btnEditar").prop("disabled", false);
      $("#btnAprovar").prop("disabled", false);
      $("#btnVoltar").prop("disabled", false);
      return false;

    }


  }


  async reprovar() {

    $("#btnReprovar").prop("disabled", true);
    $("#btnEditar").prop("disabled", true);
    $("#btnAprovar").prop("disabled", true);
    $("#btnVoltar").prop("disabled", true);

    var r = confirm("Deseja realmente reprovar esse BIT?");

    if (r == true) {

      try {

        var textoJustificativaHistorico = $("#txtJustificativa").val();

        if (textoJustificativaHistorico == "") {

          alert("Forneça uma justificativa para reprovar o BIT!");
          $("#btnReprovar").prop("disabled", false);
          $("#btnEditar").prop("disabled", false);
          $("#btnAprovar").prop("disabled", false);
          $("#btnVoltar").prop("disabled", false);
          return false;

        }

        var today = new Date();
        var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
        var date = today.getDate() + '/' + (today.getMonth() + 1) + '/' + today.getFullYear();

        await _web.lists.getByTitle("BIT").items.getById(_idBit).update({
          statusInterno: "Reprovado",
          statusbit: "Reprovado",
        }).then(async response => {

          _web.lists
            .getByTitle("Historico")
            .items.add({
              Title: _userTitle + " reprovou o BIT as  " + date + " as " + time,
              BITId: _idBit,
              Justificativa: textoJustificativaHistorico,
              TemplateEmail: "Reprovado"

            })
            .then(response => {

              console.log("reprovou");
              jQuery("#modalReprovado").modal({ backdrop: 'static', keyboard: false });


            });

        }).catch((error: any) => {
          console.log("Error: ", error);
          //this.criarLog(error.responseJSON.detailedMessage);
        });






      } catch (ex) {
        console.log(ex);;
        // this.criarLog(ex);
      }


    } else {
      $("#btnReprovar").prop("disabled", false);
      $("#btnEditar").prop("disabled", false);
      $("#btnAprovar").prop("disabled", false);
      $("#btnVoltar").prop("disabled", false);
      return false;

    }
  }

  voltar() {

    history.back();

  }

  fecharSucesso() {

    $("#modalItens").modal('hide');
    window.location.href = `BIT.aspx`;

  }



  protected async getHistorico() {

    console.log("entrou no historico");

    let montaItem = "";

    const q: ICamlQuery = {
      ViewXml: `<View><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /><FieldRef Name='Justificativa' /></ViewFields>` +
        `<Query>` +
        `<Where><Eq><FieldRef Name="RCS" LookupId='TRUE'/><Value Type="Lookup">${_idBit}</Value></Eq></Where>` +
        `<OrderBy><FieldRef Name='ID' /></OrderBy>` +
        `</Query><RowLimit>500</RowLimit>` +
        `</View>`
    };

    sp.web.lists
      .getByTitle('Historico')
      .getItemsByCAMLQuery(q, "FieldValuesAsText")
      .then((r: any[]) => {
        {
          var resRCS = "";

          if (r.length > 0) {

            //console.log("entrou!!");

            for (var index = 0; index < r.length; index++) {

              const x = r[index];

              var titulo = x.FieldValuesAsText.Title;
              var justificativa = x.FieldValuesAsText.Justificativa;

              montaItem += `<tr>
                <td style="max-width: 100px; width: 100px">2${titulo}</td>
                <td>${justificativa}</td>
            </tr>`;

            }

            $("#conteudoHistorico").html(montaItem);

          }

        }
      })
      .catch(console.error);

  }




}


