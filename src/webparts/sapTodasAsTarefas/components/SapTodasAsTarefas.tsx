import * as React from 'react';
import styles from './SapTodasAsTarefas.module.scss';
import { ISapTodasAsTarefasProps } from './ISapTodasAsTarefasProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jQuery from "jquery";

//Importação relacionada a react-bootstrap-table-next    
//Import related to react-bootstrap-table-next    
import BootstrapTable from 'react-bootstrap-table-next';
//Import from @pnp/sp    
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import { Web } from "sp-pnp-js";

import paginationFactory from 'react-bootstrap-table2-paginator';
import filterFactory, { textFilter } from 'react-bootstrap-table2-filter';
import { selectFilter } from 'react-bootstrap-table2-filter';
import { numberFilter } from 'react-bootstrap-table2-filter';
import { Comparator } from 'react-bootstrap-table2-filter';

import 'react-bootstrap-table2-paginator/dist/react-bootstrap-table2-paginator.min.css';
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _grupos;
var _web;
var _filter = "";

export interface IShowEmployeeStates {
  employeeList: any[]
}

const customFilter = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const empTablecolumns = [
  {
    dataField: "Proposta.Title",
    text: "Proposta",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },
  {
    dataField: "Proposta.Numero",
    text: "Número da Proposta",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },

  {
    dataField: "GrupoSharepoint.Title",
    text: "Atribuido a",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },
  {
    dataField: "Cliente",
    text: "Cliente",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },
  {
    dataField: "Representante",
    text: "Representante",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,

  },
  {
    dataField: "DataPlanejadaTermino",
    text: "Data Planejada de Termino",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
    formatter: (rowContent, row) => {

      var dataPlanejadaTermino = new Date(row.DataPlanejadaTermino);
      var dtDataPlanejadaTermino = ("0" + dataPlanejadaTermino.getDate()).slice(-2) + '/' + ("0" + (dataPlanejadaTermino.getMonth() + 1)).slice(-2) + '/' + dataPlanejadaTermino.getFullYear();

      return (
        <>
          {dtDataPlanejadaTermino}
        </>
      )


    }
  },
  {
    dataField: "Atraso",
    text: "Atraso",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },
  {
    dataField: "",
    text: "",
    headerStyle: { "backgroundColor": "#bee5eb", "width": "180px" },
    formatter: (rowContent, row) => {

      console.log("row",row);
      var id = row.Proposta.ID;
      var urlDetalhes = `Proposta-Detalhes.aspx?PropostasID=` + id;

      return (
        <>
          <a href={urlDetalhes}><button className="btn btn-info">Exibir</button></a>&nbsp;
        </>
      )




    }
  }


]

const paginationOptions = {
  sizePerPage: 20,
  hideSizePerPage: true,
  hidePageListOnlyOnePage: true
};

/*
const priceFilter = textFilter({
  placeholder: 'My Custom PlaceHolder',  // custom the input placeholder
});
*/

export default class SapTodasAsTarefas extends React.Component<ISapTodasAsTarefasProps, IShowEmployeeStates> {


  constructor(props: ISapTodasAsTarefasProps) {
    super(props);
    this.state = {
      employeeList: []
    }
  }


  public async componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);

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

              if (i == resultData.d.results.length - 1) {

                _filter += `(Title eq '${resultData.d.results[i].Title}')`

              } else {

                _filter += `(Title eq '${resultData.d.results[i].Title}') or `

              }

            }

            console.log("filter", _filter);

          }

        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(textStatus);
        }

      })

      console.log("grupos", grupos);
      //_grupos = grupos;
    })

    var reactHandlerRepresentante = this;

    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$orderby=%20ID%20desc&$select=ID,Title,Proposta/ID,Proposta/Title,Proposta/Numero,GrupoSharepoint/Title,DataPlanejadaTermino,Atraso,Cliente,Representante&$expand=Proposta,GrupoSharepoint&$filter=(Status eq 'Em análise') and (${_filter})`;
    console.log("url", url)

    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        jQuery('#txtCountProposta').html(resultData.d.results.length);
        reactHandlerRepresentante.setState({
          employeeList: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


  }




  public render(): React.ReactElement<ISapTodasAsTarefasProps> {
    return (


        <><p>Resultado: <span className="text-info" id="txtCountProposta"></span> proposta(s) encontrada(s)</p>
          <div className={styles.container}>
            <BootstrapTable wrapperClasses="table-responsive" bootstrap4 responsive condensed hover={true} className="gridTodosItens" id="gridTodosItens" keyField='id' data={this.state.employeeList} columns={empTablecolumns} headerClasses="header-class" pagination={paginationFactory(paginationOptions)} filter={filterFactory()} />
          </div></>


    );
  }
}
