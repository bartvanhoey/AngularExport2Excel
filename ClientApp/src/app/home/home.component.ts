import { Component, Inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import Core = require("@angular/core");


@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
})
export class HomeComponent {

  constructor(private http: HttpClient, @Inject('BASE_URL') private baseUrl: string) {

  }


  exportToExcel() {
    console.log('Export excel');

    this.http.get(this.baseUrl + 'api/ExportToExcel/GetExcelFile', {
      responseType: "blob"
    }).subscribe(result => {
      this.exportExcel(result);

      console.log(result);
    }, error => console.error(error));
  }

  exportExcel(content) {
    const blob = new Blob([content], { type: 'application/octet-stream' });
    const aLink = window.document.createElement('a');
    aLink.href = window.URL.createObjectURL(blob);
    aLink.download = 'ILikeBelgianBeers.xlsx';
    document.body.appendChild(aLink);
    aLink.click();
    document.body.removeChild(aLink);
  }
}
