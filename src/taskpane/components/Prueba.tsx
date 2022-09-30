import React from "react";
import styled from "styled-components";
// import Modal from "./Modal";

export default function Prueba() {
  return (
    <div>
      <ContenedorBotones>
        <Boton>Modal 1</Boton>
      </ContenedorBotones>
      {/* <Modal title={"Agrega una Nueva Persona"}>
        <Contenido>
          <h1>Nueva Persona</h1>
          <p>Aqui puede agregar una nueva persona</p>
        </Contenido>
      </Modal> */}
    </div>
  );
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

// const Contenido = styled.div`
//   display: flex;
//   flex-direction: column;
//   h1 {
//     font-size: 20px;
//     font-weight: 700;
//     margin-botton: 10px;
//   }
//   p {
//     font-size: 18px;
//     margin-bottom: 20px;
//   }
//   img {
//     width: 100%;
//     vertical-align: top;
//     border-radius: 3px;
//   }
// `;
