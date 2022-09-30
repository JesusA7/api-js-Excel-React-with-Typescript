import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import Modal from "./Modal";
import styled from "styled-components";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}
export interface AppState {
  listItems: HeroListItem[];
  person: InterfacePerson;
  showModal: boolean;
  showModalMessage: boolean;
}
export interface InterfacePerson {
  name: string;
  age: number;
  sex: string;
  income: number;
}
const personDefault: InterfacePerson = {
  name: "",
  age: 0,
  sex: "Masculino",
  income: 0.0,
};
export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      person: personDefault,
      showModal: false,
      showModalMessage: false,
    };
  }
  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }
  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const cl = context.workbook.getSelectedRange();

        // Read the range address
        cl.load("address");

        // Update the fill color
        cl.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${cl.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  createData = async () => {
    try {
      await Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        const ws = wb.worksheets.getActiveWorksheet();
        const cl = ws.getRange("A1:D5");
        const encabezadoDatos: Array<string> = ["Nombre", "Edad", "Sexo", "Ingresos"];
        const person1: Array<string | Number> = ["Jesús", 29, "Masculino", 3000];
        const person2: Array<string | Number> = ["Luis", 26, "Masculino", 1300];
        const person3: Array<string | Number> = ["Nico", 32, "Masculino", 2390];
        const person4: Array<string | Number> = ["Reyna", 40, "Femenino", 5500];
        cl.values = [encabezadoDatos, person1, person2, person3, person4];
        const rangeUltimo: Excel.Range = cl.getRangeEdge(Excel.KeyboardDirection.down);
        await ctx.sync();

        const tb: Excel.Table = wb.tables.add(cl, true);

        tb.name = "Tabla_de_Personas";
        const col = tb.columns.getItem("Ingresos");
        const rg = col.getDataBodyRange();
        rg.load("numberFormat");
        await ctx.sync();
        // eslint-disable-next-line office-addins/call-sync-after-load, office-addins/call-sync-before-read
        let arr = rg.numberFormat;
        arr.forEach((e) => {
          e.forEach((_e, i, a) => {
            a[i] = "#,###.00";
          });
        });
        // eslint-disable-next-line office-addins/call-sync-after-load
        rg.numberFormat = arr;
        // eslint-disable-next-line office-addins/no-empty-load
        rangeUltimo.load();
        await ctx.sync();
        rangeUltimo.select();
        // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
        console.log(rangeUltimo.address);
      });
    } catch (error) {
      console.error(error);
    }
  };
  deletePersonTable = async () => {
    try {
      await Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        const tb = wb.tables.getItemOrNullObject("Tabla_de_Personas");
        tb.delete();
        await ctx.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };
  createChartAboutPerson = async () => {
    await Excel.run(async (ctx) => {
      let tb = ctx.workbook.tables.getItemOrNullObject("Tabla_de_Personas");
      await ctx.sync();
      // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
      if (tb.isNullObject) {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const cl = ws.getRange("A1:D5");
        const encabezadoDatos: Array<string> = ["Nombre", "Edad", "Sexo", "Ingresos"];
        const person1: Array<string | Number> = ["Jesús", 29, "Masculino", 3000];
        const person2: Array<string | Number> = ["Luis", 26, "Masculino", 1300];
        const person3: Array<string | Number> = ["Nico", 32, "Masculino", 2390];
        const person4: Array<string | Number> = ["Reyna", 40, "Femenino", 5500];
        cl.values = [encabezadoDatos, person1, person2, person3, person4];
        tb = ctx.workbook.tables.add(cl, true);
        tb.name = "Tabla_de_Personas";
      }
      const body = tb.getDataBodyRange();
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const chart = ws.charts.add(Excel.ChartType.barClustered, body);
      chart.title.text = "Grafico de Personas";
      await ctx.sync();
    });
  };
  createChartInNewSheet = async () => {
    await Excel.run(async (ctx) => {
      const wsActive = ctx.workbook.worksheets.getActiveWorksheet();
      const wsCharts = ctx.workbook.worksheets.add("Charts");
      const tb = wsActive.tables.getItem("Tabla_de_Personas");
      const chart = wsCharts.charts.add(Excel.ChartType.barClustered, tb.getDataBodyRange());
      chart.title.text = "Gráfico de Personas";
      chart.name = "Grafico_de_Personas";
      chart.setPosition("A1", "F15");
      const chart2 = wsCharts.charts.add(Excel.ChartType.pie, tb.getDataBodyRange());
      chart2.title.text = "Gráfico de Personas 2";
      chart2.name = "Grafico_de_Personas_2";
      chart2.setPosition("A16", "F30");
      const chart3 = wsCharts.charts.add(Excel.ChartType.area, tb.getDataBodyRange());
      chart3.title.text = "Gráfico de Personas 3";
      chart3.name = "Grafico_de_Personas_3";
      chart3.setPosition("G1", "L15");
      const chart4 = wsCharts.charts.add(Excel.ChartType.doughnut, tb.getDataBodyRange());
      chart4.title.text = "Gráfico de Personas 4";
      chart4.name = "Grafico_de_Personas_4";
      chart4.setPosition("G16", "L30");
      await ctx.sync();
    });
  };
  enterNewPerson = async (person: InterfacePerson) => {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const tb = ws.tables.getItemOrNullObject("Tabla_de_Personas");
      await ctx.sync();
      // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
      if (!tb.isNullObject) {
        // tb.rows.add(-1, [["Grijalba", 25, "Masculino", 5500]]);
        tb.rows.add(-1, [[person.name, person.age, person.sex, person.income]]);
      }
      await ctx.sync();
    });
  };
  createNewSheet = async () => {
    await Excel.run(async (ctx) => {
      let ws: Excel.Worksheet = ctx.workbook.worksheets.getItemOrNullObject("Pivot_Table");
      const tb = ctx.workbook.tables.getItem("Tabla_de_Personas");
      await ctx.sync();
      // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
      if (ws.isNullObject) {
        ws = ctx.workbook.worksheets.add("Pivot_Table");
      }
      ws.activate();
      ws.showGridlines = false;
      const pt = ctx.workbook.pivotTables.add("PersonTable", tb, ws.getRange("B2"));
      pt.name = "Tabla_Dinamica_de_Personas";
      pt.rowHierarchies.add(pt.hierarchies.getItem("nombre"));
      const dataH1 = pt.dataHierarchies.add(pt.hierarchies.getItem("Ingresos"));
      dataH1.numberFormat = "#,##0.00";
      dataH1.name = " Ingresos";
      const chart = ws.charts.add(Excel.ChartType.doughnut, pt.layout.getRange());
      chart.title.text = "Gráfico de Dona";
      chart.showAllFieldButtons = false;
      await ctx.sync();
    });
  };
  handleChangeInput = (ev: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    this.setState({ person: { ...this.state.person, [ev.target.name]: ev.target.value } });
  };
  eventActivateSheet = async () => {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.onActivated.add(() => {
        return Excel.run(async (context: Excel.RequestContext) => {
          const rg = context.workbook.getActiveCell();
          rg.values = [["HAHAHA"]];
          // console.log(evt.worksheetId);
          await context.sync();
        });
      });
      await ctx.sync();
    });
  };
  tryCatch = (callback: any) => {
    try {
      callback();
    } catch (e) {
      console.error(e);
    }
  };
  handleShowModal = (evt: boolean) => {
    this.setState({ showModal: evt });
  };
  handleShowModalMessage = (evt: boolean) => {
    this.setState({ showModalMessage: evt });
  };
  handleSubmit = (evt: React.FormEvent<HTMLFormElement>) => {
    evt.preventDefault();
    this.enterNewPerson(this.state.person);
    this.clearPerson();
    this.setState({ showModalMessage: true });
    // this.handleShowModal(false);
  };
  clearPerson = () => {
    this.setState({ person: personDefault });
  };
  render() {
    const { title, isOfficeInitialized } = this.props;
    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }
    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Proyecto Inicial" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
        </HeroList>
        <ContenedorBotones>
          <Boton onClick={this.createData}>Crear Tabla de Personas</Boton>
          <Boton onClick={this.deletePersonTable}>Eliminar Tabla de Personas</Boton>
          <Boton onClick={() => this.tryCatch(this.createChartAboutPerson)}>
            Crear Gráfico de la Tabla de Personas
          </Boton>
          <Boton onClick={() => this.tryCatch(this.enterNewPerson(this.state.person))}>Agregar a Roger</Boton>
          <Boton
            onClick={() => {
              this.setState({ person: { name: "Piero", age: 22, sex: "Masculino", income: 2500 } });
            }}
          >
            Agregar a Nueva Persona
          </Boton>
          <Boton
            onClick={() => {
              this.tryCatch(this.createNewSheet());
            }}
          >
            Agregar Hoja de Tabla Dinamica
          </Boton>
          <Boton
            onClick={() => {
              this.handleShowModal(true);
            }}
          >
            Modal 1
          </Boton>
          <Boton
            onClick={() => {
              this.tryCatch(this.eventActivateSheet());
            }}
          >
            Agregar Evento en una hoja
          </Boton>
          <Boton
            onClick={() => {
              this.tryCatch(this.createChartInNewSheet());
            }}
          >
            Agregar gráfico en nueva hoja
          </Boton>
        </ContenedorBotones>
        <Modal title={"Agregar Empleado"} showModal={this.state.showModal} setShowModal={this.handleShowModal}>
          <Contenido>
            <form action="submit" onSubmit={this.handleSubmit}>
              <label htmlFor="name">Nombre:</label>
              <input
                type="text"
                name="name"
                id="name"
                required
                placeholder="Ingresa tu nombre completo"
                value={this.state.person.name}
                onChange={(ev) => {
                  this.handleChangeInput(ev);
                }}
              />
              <label htmlFor="age">Edad:</label>
              <input
                type="number"
                name="age"
                id="age"
                min="18"
                max="75"
                placeholder="Ingresa tu edad"
                value={this.state.person.age}
                onChange={(ev) => {
                  this.handleChangeInput(ev);
                }}
              />
              <label htmlFor="sex">Género</label>
              <select
                name="sex"
                id="sex"
                placeholder="Ingresa tu Género"
                value={this.state.person.sex}
                onChange={(ev) => {
                  this.handleChangeInput(ev);
                }}
              >
                <optgroup label="Genero">
                  <option value="Masculino">Masculino</option>
                  <option value="Femenino">Femenino</option>
                </optgroup>
              </select>
              <label htmlFor="income">Ingresos:</label>
              <input
                type="number"
                name="income"
                id="income"
                step="0.01"
                min="0.00"
                max="999999.99"
                placeholder="Ingresa tu ingreso mensual"
                value={this.state.person.income}
                onChange={(ev) => {
                  this.handleChangeInput(ev);
                }}
              />
              <input type="submit" value="Aceptar" className="btn_agregar" />
            </form>
          </Contenido>
        </Modal>
        <Modal title={"Mensaje"} showModal={this.state.showModalMessage} setShowModal={this.handleShowModalMessage}>
          <h1>Exito</h1>
          <p>Se ha guardado correctamente la nueva persona</p>
        </Modal>
      </div>
    );
  }
}
const ContenedorBotones = styled.div`
  padding: 40px;
  display: flex;
  flex-wrap: wrap;
  justify-content: center;
  gap: 20px;
`;
const Boton = styled.button`
  display: block;
  padding: 10px 30px;
  border-radius: 100px;
  color: #fff;
  border: none;
  background: #1766dc;
  cursor: pointer;
  font-family: "Roboto", sans-serif;
  font-weight: 500;
  transition: 0.3s ease all;
  &:hover {
    background: #0066ff;
  }
`;
const Contenido = styled.div`
  display: flex;
  width: 100%;
  justify-content: center;
  h1 {
    font-size: 20px;
    font-weight: 700;
    margin-botton: 10px;
  }
  p {
    font-size: 18px;
    margin-bottom: 20px;
  }
  img {
    width: 100%;
    vertical-align: top;
    border-radius: 3px;
  }
  form {
    box-sizing: border-box;
    padding: 0px;
    margin: 0px;
    width: 100%;
  }
  input {
    box-sizing: border-box;
    display: block;
    width: 100%;
    height: 28px;
  }
  select {
    width: 100%;
    height: 28px;
    display: block;
  }
  .btn_agregar {
    height: 35px;
    margin-top: 15px;
    padding: 10px 30px;
    border-radius: 5px;
    color: #fff;
    border: none;
    background: #1766dc;
    cursor: pointer;
    font-family: "Roboto", sans-serif;
    font-weight: 500;
    transition: 0.3s ease all;
    &:hover {
      background: #0066ff;
    }
  }
`;
