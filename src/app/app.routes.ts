import { Routes } from '@angular/router';
import { ProductList } from './product-list/product-list';
import { ProductDetail } from './product-detail/product-detail';

export const routes: Routes = [
  { path: '', component: ProductList, title: 'Catalog' },
  { path: 'product/:id', component: ProductDetail },
  { path: '**', redirectTo: '' },
];
