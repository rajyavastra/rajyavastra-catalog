import { HttpClient } from '@angular/common/http';
import { Injectable, inject } from '@angular/core';
import { shareReplay } from 'rxjs';

export interface Product {
  id: string;
  name: string;
  category: string;
  fabric: string;
  price: number;
  description: string;
  inStock: boolean;
  images: string[];
}

@Injectable({ providedIn: 'root' })
export class ProductService {
  private http = inject(HttpClient);
  // Relative URL so the app also works when hosted in a sub-folder.
  readonly products$ = this.http.get<Product[]>('assets/products.json').pipe(shareReplay(1));
}
