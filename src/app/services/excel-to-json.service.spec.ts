import { TestBed, inject } from '@angular/core/testing';

import { ExcelToJsonService } from './excel-to-json.service';

describe('ExcelToJsonService', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [ExcelToJsonService]
    });
  });

  it('should be created', inject([ExcelToJsonService], (service: ExcelToJsonService) => {
    expect(service).toBeTruthy();
  }));
});
