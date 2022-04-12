import * as React from 'react';
import styles from './BitAprovacaoEngenharia.module.scss';
import { IBitAprovacaoEngenhariaProps } from './IBitAprovacaoEngenhariaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';
import * as $ from "jquery";
import * as jQuery from "jquery";
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import "bootstrap";
import * as Moment from 'moment';
import MontaPaginacao from "../../../../js/main.js";
import { allowOverscrollOnElement } from 'office-ui-fabric-react';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'BitTodosItensWebPartStrings';
import { Web } from "sp-pnp-js";

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");

require("../../../../css/jPages.css");

require("../../../../js/jquery-1.8.2.min.js");
require("../../../../js/js.js");
require("../../../../js/jPages.js");
require("../../../../js/highlight.pack.js");
require("../../../../js/tabifier.js");

var _statusBIT;
var _aprovadorEngenharia = false;
var _arrAprovadorEngenharia = [];
var _userTitle;
var _web;
var _res = "";

export interface IReactGetItemsState {
  items: [
    {
      "BITNumero": "",
      "Title": "",
      "Status": "",
      "Produto": { "Title": "" },
      "Cliente": { "Title": "" },
      "Aplicacao": { "Title": "" },
      "Segmento": "",
      "Vers_x00e3_o_x0020_BIT": "",
      "Author": { "Title": "" },
      "Created": "",
    }],
}

export default class BitAprovacaoEngenharia extends React.Component<IBitAprovacaoEngenhariaProps, {}> {


  public constructor(props: IBitAprovacaoEngenhariaProps, state: IReactGetItemsState) {
    super(props);
    this.state = {
      items: [
        {
          "BITNumero": "",
          "Title": "",
          "Status": "",
          "Produto": { "Title": "" },
          "Cliente": { "Title": "" },
          "Aplicacao": { "Title": "" },
          "Segmento": "",
          "Vers_x00e3_o_x0020_BIT": "",
          "Author": { "Title": "" },
          "Created": "",
        }
      ],
    };
  }


  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

    $("#conteudo_grid").hide();

    document
      .getElementById("txtPalavra")
      .addEventListener("keyup", (e: Event) => this.Buscar());


    var loc = window.location.href;
    var path = loc.substr(0, loc.lastIndexOf('/') + 1);
    var url;


    _web.currentUser.get().then(f => {

      _userTitle = f.Title;
      console.log("_userTitle1", _userTitle);

    }).catch((error: any) => {
      console.log(error);

    });


    url = `${this.props.siteurl}/_api/web/lists/getbytitle('BIT')/items?$expand=Produto,Cliente,Aplicacao,Author,Aprovador_x0020_Engenharia&$select=ID,BITNumero,Title,Status,Produto/Title,Cliente/Title,Aplicacao/Title,Segmento,Vers_x00e3_o_x0020_BIT,Author/Title,Created,Aprovador_x0020_Engenharia/Title&$orderby= ID desc&$filter=statusInterno eq 'Aguardando Aprovação Engenharia'`;

    console.log("url", url);

    jquery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        var res = "";

        //var resultado = resultData.d.results;

        $("#txtCountBIT").html(resultData.d.results.length);

        if (resultData.d.results.length > 0) {

          $("#conteudo_grid").show();

          for (var i = 0; i < resultData.d.results.length; i++) {

            var txtProduto = "";
            var txtCliente = "";
            var txtAplicacao = "";

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

              }

            }

            var id = resultData.d.results[i].ID;
            var bitNumero = resultData.d.results[i].BITNumero;
            var title = resultData.d.results[i].Title;
            var status = resultData.d.results[i].Status;
            var segmento = resultData.d.results[i].Segmento;
            var versaoBIT = resultData.d.results[i].Vers_x00e3_o_x0020_BIT;
            var author = resultData.d.results[i].Author.Title;

            var created = resultData.d.results[i].Created;
            created = Moment(resultData.d.results[i].Created).format('DD/MM/YYYY');

            if (bitNumero == null) bitNumero = "";
            if (title == null) title = "";
            if (status == null) status = "";
            if (segmento == null) segmento = "";
            if (versaoBIT == null) versaoBIT = "";
            if (author == null) author = "";

            console.log("_userTitle2", _userTitle);
            console.log("_arrAprovadorEngenharia", _arrAprovadorEngenharia);

            if (_arrAprovadorEngenharia.indexOf(_userTitle) !== -1) {

              res += `<tr>
                <td scope="col">${bitNumero}</td>
                <td scope="col">${title}</td>
                <td scope="col">${status}</td>
                <td scope="col">${txtProduto}</td>
                <td scope="col">${txtCliente}</td>
                <td scope="col">${txtAplicacao}</td>
                <td scope="col">${segmento}</td>
                <td scope="col">${versaoBIT}</td>
                <td scope="col">${author}</td>
                <td scope="col">${created}</td>
                <td scope="col">
                <div style="width: 150px;">
                <button onclick="location.href='${path}Detalhes-BIT.aspx?idBIT=${id}';" class="btn btn-info">Exibir</button>
                </div>
                </td>
                </tr>`;

            }

          }

          $("#itemContainer").html(res);
          MontaPaginacao("itemContainer", "holder", 10);

        } else {

        }





      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });




  }

  public render(): React.ReactElement<IBitAprovacaoEngenhariaProps> {
    return (


      <>

        <input id="txtPalavra" type="text" style={{ "width": "300px" }} className="form-control rounded mr-2" placeholder="forneça um termo para busca" aria-label=""
          aria-describedby="button-addon2"></input>

        <br></br>

        <div id="conteudo_grid">
          <div className="table-responsive">
            <table className="table table-hover">
              <thead className="thead-light ">
                <tr>
                  <th scope="col">BIT Número</th>
                  <th scope="col">Título</th>
                  <th scope="col">Status</th>
                  <th scope="col">Produto</th>
                  <th scope="col">Cliente</th>
                  <th scope="col">Aplicação</th>
                  <th scope="col">Segmento</th>
                  <th scope="col">Versão BIT</th>
                  <th scope="col">Criado por</th>
                  <th scope="col">Criado</th>
                  <th scope="col">Ação</th>
                </tr>
              </thead>
              <tbody id="itemContainer">
              </tbody>
            </table>
          </div>
          <hr />
          <div id="holder" className="holder">
          </div>
        </div>

      </>

    );
  }

  Buscar() {

    var termo = $("#txtPalavra").val();
    console.log(termo);
    this.buscarTermo(termo);

  }

  buscarTermo(termo) {

    //$("#itemContainer").hide();
    //$("#conteudo_nenhumaEncontrada").hide();
    var url;

    var loc = window.location.href;
    var path = loc.substr(0, loc.lastIndexOf('/') + 1);

    if (termo == "") {
      console.log("em branco");
      url = `${this.props.siteurl}/_api/web/lists/getbytitle('BIT')/items?$expand=Produto,Cliente,Aplicacao,Author,Aprovador_x0020_Engenharia&$select=ID,BITNumero,Title,Status,Produto/Title,Cliente/Title,Aplicacao/Title,Segmento,Vers_x00e3_o_x0020_BIT,Author/Title,Created,Aprovador_x0020_Engenharia/Title&$orderby= ID desc&$filter=statusInterno eq 'Aguardando Aprovação Engenharia'`;
    }
    else {

      url = `${this.props.siteurl}/_api/web/lists/getbytitle('BIT')/items?$expand=Produto,Cliente,Aplicacao,Author,Aprovador_x0020_Engenharia&$select=ID,BITNumero,Title,Status,Produto/Title,Cliente/Title,Aplicacao/Title,Segmento,Vers_x00e3_o_x0020_BIT,Author/Title,Created,Aprovador_x0020_Engenharia/Title&$orderby= ID desc&$filter=((substringof('` + termo + `', Title)) or (substringof('` + termo + `', Cliente/Title))) and (statusInterno eq 'Aguardando Aprovação Engenharia')`;

    }

    console.log("url", url);

    jquery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        var res = "";

        $("#txtCountBIT").empty();
        $("#txtCountBIT").html(resultData.d.results.length);

        if (resultData.d.results.length > 0) {

          console.log();

          for (var i = 0; i < resultData.d.results.length; i++) {

            var txtProduto = "";
            var txtCliente = "";
            var txtAplicacao = "";

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

              }

            }


            var id = resultData.d.results[i].ID;
            var bitNumero = resultData.d.results[i].BITNumero;
            var title = resultData.d.results[i].Title;
            var status = resultData.d.results[i].Status;
            var segmento = resultData.d.results[i].Segmento;
            var versaoBIT = resultData.d.results[i].Vers_x00e3_o_x0020_BIT;
            var author = resultData.d.results[i].Author.Title;

            var created = resultData.d.results[i].Created;
            created = Moment(resultData.d.results[i].Created).format('DD/MM/YYYY');

            console.log("created", created);

            if (bitNumero == null) bitNumero = "";
            if (title == null) title = "";
            if (status == null) status = "";
            if (segmento == null) segmento = "";
            if (versaoBIT == null) versaoBIT = "";
            if (author == null) author = "";

            if (_arrAprovadorEngenharia.indexOf(_userTitle) !== -1) {

              res += `<tr>
            <td scope="col">${bitNumero}</td>
            <td scope="col">${title}</td>
            <td scope="col">${status}</td>
            <td scope="col">${txtProduto}</td>
            <td scope="col">${txtCliente}</td>
            <td scope="col">${txtAplicacao}</td>
            <td scope="col">${segmento}</td>
            <td scope="col">${versaoBIT}</td>
            <td scope="col">${author}</td>
            <td scope="col">${created}</td>
            <td scope="col">
            <div style="width: 150px;">
            <button onclick="location.href='${path}Detalhes-BIT.aspx?idBIT=${id}';" class="btn btn-info">Exibir</button>
            </div>
            </td>

          </tr>`

            }


          }


        }

        console.log("res", res);
        $("#itemContainer").html(res);



        MontaPaginacao("itemContainer", "holder", 10);



      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });




  }

}
