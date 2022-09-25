import React, { useState } from "react";
import { Link } from "react-router-dom";
import "./header.css";
const Navbar = () => {
  const [isOpen, setIsOpen] = useState(false);
  return (
    <div className="Navbar">
      <span className="nav-logo">Comptabilité sociale et environnementale</span>
      <div className={`nav-items ${isOpen && "open"}`}>
        
        <a href="/aboutCSE">À propos du CSE</a>
        <a href="https://gbl-inc.org/">À propos de nous</a>
        
      </div>
      <div
        className={`nav-toggle ${isOpen && "open"}`}
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="bar"></div>
      </div>
    </div>
  );
};
export default Navbar
