import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Progress from "./Progress";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  message:string
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      message: "Hello World",
    };
  }

  componentDidMount() {   
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    
    return (
      <div className="ms-welcome">
        <Header logo={require("./../../assets/logo-filled.png")} title={this.props.title} message={this.state.message} />
    </div>
    );
  }
}
