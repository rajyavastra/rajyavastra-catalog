import { Component } from '@angular/core';
import { RouterLink, RouterOutlet } from '@angular/router';
import { STORE } from './config';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, RouterLink],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  store = STORE;
  petals = [1, 2, 3, 4, 5, 6, 7];
}
