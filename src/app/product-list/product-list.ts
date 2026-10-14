import { CurrencyPipe } from '@angular/common';
import { Component, computed, inject, signal } from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { RouterLink } from '@angular/router';
import { STORE } from '../config';
import { ProductService } from '../product';

@Component({
  selector: 'app-product-list',
  imports: [RouterLink, CurrencyPipe],
  templateUrl: './product-list.html',
  styleUrl: './product-list.css',
})
export class ProductList {
  currency = STORE.currency;
  products = toSignal(inject(ProductService).products$);

  query = signal('');
  category = signal('All');

  categories = computed(() => ['All', ...new Set((this.products() ?? []).map((p) => p.category))]);

  visible = computed(() => {
    const q = this.query().trim().toLowerCase();
    const cat = this.category();
    return (this.products() ?? []).filter(
      (p) =>
        (cat === 'All' || p.category === cat) &&
        (!q || p.name.toLowerCase().includes(q) || p.description.toLowerCase().includes(q)),
    );
  });
}
