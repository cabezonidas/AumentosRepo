import * as React from 'react';
import { IAumentosProps } from './IAumentosProps';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from 'sp-pnp-js';
import { Slider } from 'office-ui-fabric-react/lib/Slider';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PrimaryButton } from "office-ui-fabric-react/lib/components/Button";
import { Dialog, DialogType } from "office-ui-fabric-react/lib/Dialog";
import { ProgressIndicator } from "office-ui-fabric-react/lib/ProgressIndicator";

export default class Aumentos extends React.Component<IAumentosProps, any> {
    constructor() {
    super();
    this.state = {
      examples: [],
      canPreview: false,
      percentage: 20,
      completition: 0,
      saving: false,
      saved: false
    }; 
  }  
  public render(): React.ReactElement<IAumentosProps> {
    return (
      <div>
        <h3>Ajustar porcentaje de aumento</h3>
        <br/>
        <Slider label='Porcentaje:' min={ 1 } max={ 100 } value={ this.state.percentage } onChange={ value => this.setState({ percentage: value }) } showValue={ true } />        
        <br/>
        <h3>Vista preliminar</h3>
        <div className={'ms-Grid'} style={{marginLeft: '20px', marginRight: '20px'}}>
          <div className={'ms-Grid-row'}>
            <Label className={'ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6'}>
              <strong>Producto</strong>              
            </Label>
            <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{textAlign: 'center'}}>
              <strong>Precio</strong>
            </Label>
            <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{textAlign: 'center'}}>
              <strong>Aumento</strong>
            </Label>
            <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{textAlign: 'center'}}>
              <strong>Redondeo</strong>
            </Label>
          </div>
        </div>
        <div style={{overflow: 'auto', height:'250px'}}>
          { 
            this.state.examples.map(product => 
              <div className={'ms-Grid'} style={{marginLeft: '20px', marginRight: '20px'}}>
                <div className={'ms-Grid-row'}>
                  <Label className={'ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6'}>{product.Title}</Label>
                  <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{color: 'red', textAlign: 'right'}}>
                    <strong>{product.Precio.toFixed(2)}</strong>
                  </Label>
                  <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{color: 'orange', textAlign: 'right'}}>
                    <strong>{this.calculateValue(product.Precio)}</strong>
                  </Label>
                  <Label className={'ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2'} style={{color: 'green', textAlign: 'right'}}>
                    <strong>{this.roundValue(product.Precio)}</strong>
                  </Label>
                </div>
              </div>
            ) 
          }
        </div>
        <br/>
        <PrimaryButton type='button' label='Guardar' onClick={this.updatePrices}/>
        <Dialog
          isOpen={ this.state.saving }
          type={ DialogType.normal }
          title='Guardando'
          subText={ !(this.state.saved && this.state.saving) ? 'No cerrar hasta que se termine el proceso...' : 'Completado!' }
          isBlocking={ true }
          containerClassName='ms-dialogMainOverride'
        >
          <ProgressIndicator percentComplete={this.state.completition} />
          <PrimaryButton label='Salir' onClick={()=> {window.location.replace('/')}} disabled={!this.state.saved} />
        </Dialog>
      </div>
    );
  }

  private componentDidMount = () => 
    pnp.sp.web.lists.getByTitle('Productos').items.top(500).orderBy('Categoria').orderBy('Title').get()
      .then(items => this.setState({examples: items, canPreview: true}));

  private updatePrices = () => {
    pnp.sp.web.lists.getByTitle("Productos").items.top(500).get().then((items: any[]) => {
        this.setState({saving: true});
        for(let i = 0; i < items.length; i++){
          pnp.sp.web.lists.getByTitle("Productos").items.getById(items[i].Id).update({
            Precio: this.roundValue(items[i].Precio)
          }).then(item => {
            this.setState({completition: (i + 1)/items.length});
            if((i + 1)/items.length == 1)
              this.setState({saved: true, completition: 1});
          });
        }
    });
  }

  private calculateValue = (precio) => {
    return (precio * (1 + (this.state.percentage / 100))).toFixed(2);
  }

  private roundValue = (precio) => {
    return Math.round(parseFloat(this.calculateValue(precio))).toFixed(2);
  }
}