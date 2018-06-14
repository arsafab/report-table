import { Component, OnInit, Input } from '@angular/core';
import { IGroup } from '../../models';

@Component({
  selector: 'app-table',
  templateUrl: './table.component.html',
  styleUrls: ['./table.component.css']
})
export class TableComponent implements OnInit {
  @Input() public data: IGroup;
  @Input() public date: Date;

  constructor() {}

  ngOnInit() {}
}
