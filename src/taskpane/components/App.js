import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import {useEffect, useState} from "react";

/* global Word, require */

const App = ({isOfficeInitialized, title}) => {
  const [listItems, setListItems] = useState([
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
  ]);
  
  const handleClick = async() => {
    return Word.run(async (context) => {
      try {
        const images = context.document.body.inlinePictures
        images.load('imageFormat')
        await context.sync()
    
        const imageData = images.items.map((image) => {
          return {
            format: image?.imageFormat,
            base64: image?.getBase64ImageSrc()
          }
        })
        await context.sync()
    
        console.log(imageData)
      } catch (error) {
        console.log('Error: ', error)
      }
    });
  };

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
        <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
            Run
          </DefaultButton>
        </HeroList>
      </div>
    );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App
