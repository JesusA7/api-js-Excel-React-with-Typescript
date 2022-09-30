import React from "react";
import styled from "styled-components";

// eslint-disable-next-line react/prop-types
const Modal = ({ title, children, showModal, setShowModal }) => {
  return (
    <>
      {showModal && (
        <Overlay>
          <ContenedorModal>
            <EncabezadoModal>
              <h3>{title}</h3>
            </EncabezadoModal>
            <BotonCerrar
              onClick={() => {
                setShowModal(false);
              }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" fill="currentColor" className="bi bi-x" viewBox="0 0 16 16">
                <path d="M4.646 4.646a.5.5 0 0 1 .708 0L8 7.293l2.646-2.647a.5.5 0 0 1 .708.708L8.707 8l2.647 2.646a.5.5 0 0 1-.708.708L8 8.707l-2.646 2.647a.5.5 0 0 1-.708-.708L7.293 8 4.646 5.354a.5.5 0 0 1 0-.708z" />
              </svg>
            </BotonCerrar>
            {children}
          </ContenedorModal>
        </Overlay>
      )}
    </>
  );
};
export default Modal;
const Overlay = styled.div`
  box-sizing: border-box;
  width: 100vw;
  height: 100vh;
  position: fixed;
  top: 0;
  left: 0;
  background-color: rgba(0, 0, 0, 0.5);
  padding: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
`;
const ContenedorModal = styled.div`
  box-sizing: border-box;
  width: 80%;
  min-width: 200px;
  min-heigth: 200px;
  background: #fff;
  position: relative;
  border-radius: 5px;
  box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;
  padding: 20px;
`;
const EncabezadoModal = styled.div`
  display: flex;
  align-items: top;
  justify-contents: space-between;
  margin-top: 0px;
  margin-bottom: 15px;
  padding-bottom: 15px;
  border-bottom: 1px solid #e8e8e8;
  h3 {
    font-weight: 500;
    font-size: 12px;
    color: #1766dc;
    margin: 0px;
  }
`;
const BotonCerrar = styled.button`
  position: absolute;
  top: 15px;
  right: 15px;
  width: 30px;
  height: 30px;
  border: none;
  background: none;
  cursor: pointer;
  transition: 0.1s ease all;
  border-radius: 5px;
  color: #1766dc;
  padding: 0px;
  &:hover {
    background: #f2f2f2;
  }

  svg {
    widht: 100%;
    height: 100%;
  }
`;
