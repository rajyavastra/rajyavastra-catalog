import { CurrencyPipe } from '@angular/common';
import { Component, computed, inject, input, signal } from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { RouterLink } from '@angular/router';
import { STORE } from '../config';
import { ProductService } from '../product';

@Component({
  selector: 'app-product-detail',
  imports: [RouterLink, CurrencyPipe],
  templateUrl: './product-detail.html',
  styleUrl: './product-detail.css',
})
export class ProductDetail {
  // Bound from the :id route parameter via withComponentInputBinding() in app.config.ts
  id = input.required<string>();
  currency = STORE.currency;
  private all = toSignal(inject(ProductService).products$);
  loaded = computed(() => this.all() !== undefined);
  product = computed(() => this.all()?.find((p) => p.id === this.id()));
  selected = signal(0);

  shareUrl = computed(() => {
    const p = this.product();
    if (!p) return '';
    const price = new Intl.NumberFormat('en-CA', { style: 'currency', currency: this.currency }).format(p.price);
    const text = `${p.name} - ${price}\n${location.href}`;
    return 'https://wa.me/?text=' + encodeURIComponent(text);
  });

  askUrl = computed(() => {
    const p = this.product();
    if (!p || !STORE.whatsappNumber) return '';
    const text = `Hi, is the ${p.name} (${p.id}) available?`;
    return `https://wa.me/${STORE.whatsappNumber}?text=` + encodeURIComponent(text);
  });
}
