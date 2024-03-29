import * as React from 'react';
import styles from './SapTodasAsTarefasAdm.module.scss';
import { ISapTodasAsTarefasAdmProps } from './ISapTodasAsTarefasAdmProps';
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
var _siteurl = "";
var _representante;

export interface IShowEmployeeStates {
  employeeList: any[]
}

const customFilter = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const empTablecolumns = [
  {
    dataField: "Proposta.Numero",
    text: "Número",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
    classes: 'text-center',
    formatter: (rowContent, row) => {
      var tarefaAntiga = row.TarefaAntiga;
      var val;
      if(tarefaAntiga == "Sim"){
        val = row.AntigoPropostaNumero;
        val = val.replace(".000000000","");
      }else{
        val = row.Proposta.Numero;
      }
      return val;
    }
  },
  {
    dataField: "Proposta.IdentificacaoOportunidade",
    text: "Oportunidade",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    classes: 'text-center',
    filter: customFilter,
    formatter: (rowContent, row) => {
      var tarefaAntiga = row.TarefaAntiga;
      var val;
      if(tarefaAntiga == "Sim"){
        val = row.AntigoPropostaOportunidade;
      }else{
        val = row.Proposta.IdentificacaoOportunidade;
      }
      return val;
    }

  },
  {
    dataField: "Created",
    text: "Data de criação",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
    classes: 'text-center',
    formatter: (rowContent, row) => {
      var dataCriacao = new Date(row.Created);
      var dtdataCriacao = ("0" + dataCriacao.getDate()).slice(-2) + '/' + ("0" + (dataCriacao.getMonth() + 1)).slice(-2) + '/' + dataCriacao.getFullYear();
      return dtdataCriacao;
    }
  },
  {
    dataField: "Cliente",
    text: "Cliente",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter
  },
  {
    dataField: "Proposta.Title",
    text: "Síntese",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
    formatter: (rowContent, row) => {
      var tarefaAntiga = row.TarefaAntiga;
      var val;
      if(tarefaAntiga == "Sim"){
        val = row.Title;
      }else{
        val = row.Proposta.Title;
      }
      return val;
    }

  },
  {
    dataField: "Proposta.DataEntregaPropostaCliente",
    text: "Data de entrega",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
    formatter: (rowContent, row) => {
      var dataEntregaPropostaCliente = new Date(row.Proposta.DataEntregaPropostaCliente);
      var dtDataEntregaPropostaCliente = ("0" + dataEntregaPropostaCliente.getDate()).slice(-2) + '/' + ("0" + (dataEntregaPropostaCliente.getMonth() + 1)).slice(-2) + '/' + dataEntregaPropostaCliente.getFullYear();
      //console.log("dtDataEntregaPropostaCliente", dtDataEntregaPropostaCliente);
      return dtDataEntregaPropostaCliente;
    }
  },
  {
    dataField: "GrupoSharepoint.Title",
    text: "Atribuido a",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,
  },
  {
    dataField: "Representante",
    text: "Representante",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilter,

  },
  /*
  {
    dataField: "DataPlanejadaTermino",
    text: "Data Planejada de Termino",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    //filter: customFilter,
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
    //filter: customFilter
  },
  */
  {
    dataField: "",
    text: "",
    headerStyle: { "backgroundColor": "#bee5eb", "width": "180px" },
    formatter: (rowContent, row) => {
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


export default class SapTodasAsTarefasAdm extends React.Component<ISapTodasAsTarefasAdmProps, IShowEmployeeStates> {

  constructor(props: ISapTodasAsTarefasAdmProps) {
    super(props);
    this.state = {
      employeeList: []
    }
  }

  public async componentDidMount() {



    _web = new Web(this.props.context.pageContext.web.absoluteUrl);




    _siteurl = this.props.siteurl;

    var reactHandler = this;

    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Tarefas')/items?$top=4999&$orderby=%20ID%20desc&$select=ID,Title,Proposta/ID,Proposta/Title,Proposta/Numero,Proposta/IdentificacaoOportunidade,Proposta/DataEntregaPropostaCliente,Proposta/ResponsavelProposta,GrupoSharepoint/Title,DataPlanejadaTermino,Atraso,Created,Author/Title,Cliente,Representante,TarefaAntiga,AntigoPropostaNumero,AntigoPropostaOportunidade&$expand=Proposta,GrupoSharepoint,Author&$filter=(Status eq 'Em análise')`;
    //console.log("url", url);

    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData",resultData);

        jQuery('#txtCountProposta').html(resultData.d.results.length);

        reactHandler.setState({
          employeeList: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


  }


  public render(): React.ReactElement<ISapTodasAsTarefasAdmProps> {
    return (

      <><p>Aprovações encontradas: <span className="text-info" id="txtCountProposta"></span></p>
        <div className={styles.container}>
          <BootstrapTable bootstrap4 responsive condensed hover={true} className="gridTodosItens" id="gridTodosItens" keyField='id' data={this.state.employeeList} columns={empTablecolumns} headerClasses="header-class" pagination={paginationFactory(paginationOptions)} filter={filterFactory()} />
        </div></>


    );
  }
}
