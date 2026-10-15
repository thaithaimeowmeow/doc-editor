import { ChangeDetectionStrategy, Component, ElementRef, effect, inject, input } from '@angular/core';

import { ICONS } from './icons';

/** Icon SVG tĩnh từ ICONS (chuỗi hằng trong code, gán trực tiếp nên không qua sanitizer của Angular). */
@Component({
  selector: 'we-icon',
  template: '',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { 'aria-hidden': 'true', class: 'we-icon' },
})
export class WeIcon {
  readonly name = input.required<string>();

  constructor() {
    const host = inject<ElementRef<HTMLElement>>(ElementRef).nativeElement;
    effect(() => {
      host.innerHTML =
        `<svg viewBox="0 0 20 20" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.4" ` +
        `stroke-linecap="round" stroke-linejoin="round">${ICONS[this.name()] ?? ''}</svg>`;
    });
  }
}
