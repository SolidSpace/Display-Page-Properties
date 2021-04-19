import * as React from 'react';
export interface ITestComponentProps {
  text:string;
}

export interface ITestComponentState {}

export default class TestComponent extends React.Component<ITestComponentProps, ITestComponentState> {
  public render(): React.ReactElement<ITestComponentProps> {
    return (
      <div>
        <div>{this.props.text}</div>
      </div>
    );
  }
}
