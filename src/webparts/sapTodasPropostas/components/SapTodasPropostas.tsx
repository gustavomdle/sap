import * as React from 'react';
import styles from './SapTodasPropostas.module.scss';
import { ISapTodasPropostasProps } from './ISapTodasPropostasProps';
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

export interface IShowEmployeeStates {
  employeeList: any[]
}

const selectOptions = {
  0: 'Aprovado',
  1: 'Em análise',
  2: 'Encerrada pelo Sistema',
  3: 'Não vencedora',
  4: 'Proposta Enviada',
  5: 'Reprovada',
  6: 'Vencedora',
};


const customFilterID = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterNumero = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterIDOportunidade = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterIDSintese = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterCliente = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterRepresentante = textFilter({
  placeholder: ' ',  // custom the input placeholder
});

const customFilterStatus = selectFilter({
  placeholder: 'Selecione',  // custom the input placeholder
});





const empTablecolumns = [
  {
    dataField: "ID",
    text: "ID",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterID
  },
  {
    dataField: "Numero",
    text: "Número",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterNumero
  },
  {
    dataField: "IdentificacaoOportunidade",
    text: "ID Oportunidade",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterIDOportunidade
  },
  {
    dataField: "Title",
    text: "Síntese",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterIDSintese
  },
  {
    dataField: "Cliente.Title",
    text: "Cliente",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterCliente
  },
  {
    dataField: "Representante.Title",
    text: "Representante",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: customFilterRepresentante
  },
  {
    dataField: "Status",
    text: "Status",
    headerStyle: { backgroundColor: '#bee5eb' },
    sort: true,
    filter: selectFilter({
      options: selectOptions
    })
  },
  {
    dataField: "",
    text: "",
    headerStyle: { "backgroundColor": "#bee5eb", "width": "180px" },
    formatter: (rowContent, row) => {
      var id = row.ID;
      var status = row.Status
      var urlDetalhes = `Proposta-Detalhes.aspx?PropostasID=` + id;
      var urlEditar = `Propostas-SAP-Editar.aspx?PropostasID=` + id;

      if (status == "Em análise") {

        console.log("_grupos",_grupos);

        if (_grupos.indexOf("Representante") !== -1) {
          return (
            <>
              <a href={urlDetalhes}><button className="btn btn-info">Exibir</button></a>&nbsp;
              <a href={urlEditar}><button className="btn btn-danger">Editar</button></a>
            </>
          )
        } else {

          return (
            <>
              <a href={urlDetalhes}><button className="btn btn-info">Exibir</button></a>&nbsp;
            </>
          )
        }

      } else {
        return (
          <>
            <a href={urlDetalhes}><button className="btn btn-info">Exibir</button></a>&nbsp;
          </>
        )
      }


    }
  }

];

const paginationOptions = {
  sizePerPage: 5,
  hideSizePerPage: true,
  hidePageListOnlyOnePage: true
};

const priceFilter = textFilter({
  placeholder: 'My Custom PlaceHolder',  // custom the input placeholder
});



export default class SapTodasPropostas extends React.Component<ISapTodasPropostasProps, IShowEmployeeStates> {



  constructor(props: ISapTodasPropostasProps) {
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

    var reactHandlerRepresentante = this;

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('PropostasSAP')/items?$top=4999&$orderby= ID desc&$select=ID,Title,Numero,IdentificacaoOportunidade,Title,Cliente/Title,Representante/Title,Status&$expand=Cliente,Representante`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        // console.log("resultData.d.results", resultData.d.results)
        reactHandlerRepresentante.setState({
          employeeList: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });


  }


  Buscar(): any {
    throw new Error('Method not implemented.');
  }


  public render(): React.ReactElement<ISapTodasPropostasProps> {




    return (


      <div className={styles.container}>
        <BootstrapTable bootstrap4 responsive condensed hover={true} className="gridTodosItens" id="gridTodosItens" keyField='id' data={this.state.employeeList} columns={empTablecolumns} headerClasses="header-class" pagination={paginationFactory(paginationOptions)} filter={filterFactory(priceFilter)} />
      </div>


    );
  }



}
