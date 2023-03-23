import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';

import { AppComponent } from './app.component';
import { HelloComponent } from './hello.component';
import { GeneratorComponent } from './modules/generator.component';
import { MaterialModule } from './material.module';
import { DragDropDirective } from './directives/drag-drop.directive';
import { PdfViewerModule } from 'ng2-pdf-viewer';

@NgModule({
  declarations: [AppComponent, HelloComponent, GeneratorComponent, DragDropDirective],
  imports: [BrowserModule, FormsModule, ReactiveFormsModule, MaterialModule,PdfViewerModule],
  bootstrap: [AppComponent]
})
export class AppModule { }
