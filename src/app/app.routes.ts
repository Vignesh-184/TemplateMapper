import { Routes } from '@angular/router';
import { UploadComponent } from './modules/upload/upload.component';
import { MappingComponent } from './modules/mapping/mapping.component';
import { PreviewComponent } from './modules/preview/preview.component';

export const routes: Routes = [
  { path: '', redirectTo: 'upload', pathMatch: 'full' },
  { path: 'upload', component: UploadComponent },
  { path: 'mapping', component: MappingComponent },
  { path: 'preview', component: PreviewComponent }
];
