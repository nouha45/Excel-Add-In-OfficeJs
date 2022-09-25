import * as React from "react";
import Navbar from "./Navbar";
import Footer from "./Footer";
import Home from "./Home";


/* global console, Excel, require */

function App() {
    return (
      
      <div className="ms-welcome">
        <Navbar/>
         {/* <BrowserRouter>
         <Navbar/>
        <Routes> 
        <Route path='/' exact element={<Home/>}/>
          <Route path='/aboutus' exact element={<AboutUsPage/>}/> 

       </Routes>
        </BrowserRouter>  */}
         {/* <AboutUsPage/>  */}
        <Home/>
        <Footer/>
       
      
      </div>
    );
  }
  export default App;


