import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { environment } from 'src/environments/environment';

@Injectable({
  providedIn: 'root',
})
export class ExportService {
  constructor(private http: HttpClient) {}

  public saveFormAsExcel(form: any): void {
    this.exportFormToExcel(form).subscribe((response) => {
      const url = window.URL.createObjectURL(response);
      window.open(url);
    });
  }

  public saveDatatableAsExcel(datatable: any[]): void {
    this.exportDatatableToExcel(datatable).subscribe((response) => {
      const url = window.URL.createObjectURL(response);
      window.open(url);
    });
  }

  public saveDatatablesAsExcel(datatables: any[]): void {
    this.exportDatatablesToExcel(datatables).subscribe((response) => {
      const url = window.URL.createObjectURL(response);
      window.open(url);
    });
  }

  public saveFormAsPdf(form: any): void {
    this.exportFormToPdf(form).subscribe((response) => {
      const file = new Blob([response], { type: 'application/pdf' });
      const fileURL = URL.createObjectURL(file);
      window.open(fileURL);
    });
  }

  public saveDatatableAsPdf(datatable: any[]): void {
    this.exportDatatableToPdf(datatable).subscribe((response) => {
      const file = new Blob([response], { type: 'application/pdf' });
      const fileURL = URL.createObjectURL(file);
      window.open(fileURL);
    });
  }

  private exportFormToExcel(form: any): Observable<Blob> {
    const url = environment.apiURL + '/export/form-to-excel';
    return this.http.post(url, form, {
      responseType: 'blob',
    });
  }

  private exportDatatableToExcel(datatable: any[]): Observable<Blob> {
    const url = environment.apiURL + '/export/datatable-to-excel';
    return this.http.post(url, datatable, {
      responseType: 'blob',
    });
  }

  private exportDatatablesToExcel(datatables: any[]): Observable<Blob> {
    const url = environment.apiURL + '/export/datatables-to-excel';
    return this.http.post(url, datatables, {
      responseType: 'blob',
    });
  }

  private exportFormToPdf(form: any): Observable<Blob> {
    const url = environment.apiURL + '/export/form-to-pdf';
    let headers = new HttpHeaders();
    headers = headers.set('Accept', 'application/pdf');
    return this.http.post(url, form, {
      headers: headers,
      responseType: 'blob',
    });
  }

  private exportDatatableToPdf(datatable: any[]): Observable<Blob> {
    const url = environment.apiURL + '/export/datatable-to-pdf';
    let headers = new HttpHeaders();
    headers = headers.set('Accept', 'application/pdf');
    return this.http.post(url, datatable, {
      headers: headers,
      responseType: 'blob',
    });
  }
}
