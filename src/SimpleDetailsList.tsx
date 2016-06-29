import * as React from 'react';
import {
  ColumnActionsMode,
  CommandBar,
  ConstrainMode,
  ContextualMenu,
  DetailsList,
  DetailsListLayoutMode as LayoutMode,
  DirectionalHint,
  IColumn,
  IContextualMenuItem,
  IContextualMenuProps,
  Link,
  TextField,
  Toggle,
  buildColumns
} from 'office-ui-fabric-react';
import { SelectionMode } from 'office-ui-fabric-react/lib/utilities/selection/interfaces';

const PAGING_SIZE = 1000;
const PAGING_DELAY = 5000;
const ITEMS_COUNT = 5000;

let _items;

export interface IDetailsListBasicExampleState {
  items?: any[];
  layoutMode?: LayoutMode;
  constrainMode?: ConstrainMode;
  selectionMode?: SelectionMode;
  canResizeColumns?: boolean;
  columns?: IColumn[];
  sortedColumnKey?: string;
  isSortedDescending?: boolean;
  contextualMenuProps?: IContextualMenuProps;
  isLazyLoaded?: boolean;
  isHeaderVisible?: boolean;
  isGridVisible?: boolean;
}

export class SimpleDetailsList extends React.Component<any, IDetailsListBasicExampleState> {
  private _isFetchingItems: boolean;

  constructor() {
    super();

    if (!_items) {
      _items = this.createListItems(ITEMS_COUNT);
    }

    this._onToggleResizing = this._onToggleResizing.bind(this);
    this._onToggleLazyLoad = this._onToggleLazyLoad.bind(this);
    this._onLayoutChanged = this._onLayoutChanged.bind(this);
    this._onConstrainModeChanged = this._onConstrainModeChanged.bind(this);
    this._onSelectionChanged = this._onSelectionChanged.bind(this);
    this._onColumnClick = this._onColumnClick.bind(this);
    this._onContextualMenuDismissed = this._onContextualMenuDismissed.bind(this);

    this.state = {
      items: _items, // createListItems(100).concat(new Array(9900)), // _items,
      layoutMode: LayoutMode.justified,
      constrainMode: ConstrainMode.horizontalConstrained,
      selectionMode: SelectionMode.multiple,
      canResizeColumns: true,
      columns: this._buildColumns(true, this._onColumnClick, ''),
      contextualMenuProps: null,
      sortedColumnKey: 'name',
      isSortedDescending: false,
      isLazyLoaded: false,
      isHeaderVisible: true,
      isGridVisible: false,
    };
  }

  public render() {
    let {
      items,
      layoutMode,
      constrainMode,
      selectionMode,
      columns,
      contextualMenuProps,
      isHeaderVisible,
      isGridVisible
    } = this.state;

    return (
      <div className='ms-DetailsListBasicExample'>
        <Toggle
          isToggled={ isGridVisible }
          onChanged={ isToggled => this.setState({ isGridVisible: isToggled }) }
          label='Grid visible'
          onText='On'
          offText='Off' 
          />

        <CommandBar items={ this._getCommandItems() } />

        <div style={ isGridVisible ? null : { display: "none" } }>
          <DetailsList
            setKey='items'
            items={ items }
            columns={ columns }
            layoutMode={ layoutMode }
            isHeaderVisible={ isHeaderVisible }
            selectionMode={ selectionMode }
            constrainMode={ constrainMode }
            onItemInvoked={ this._onItemInvoked }
            ariaLabelForListHeader='Column headers. Use menus to perform column operations like sort and filter'
            ariaLabelForSelectAllCheckbox='Toggle selection for all items'
            onRenderMissingItem={ (index) => {
              this._onDataMiss(index);
              return null;
            } }
            />
          </div>

        { contextualMenuProps && (
          <ContextualMenu { ...contextualMenuProps } />
        ) }
      </div>
    );
  }

  private _onDataMiss(index) {
    index = Math.floor(index / PAGING_SIZE) * PAGING_SIZE;

    if (!this._isFetchingItems) {

      this._isFetchingItems = true;

      setTimeout(() => {
        this._isFetchingItems = false;
        let itemsCopy = [].concat(this.state.items);

        itemsCopy.splice.apply(itemsCopy, [index, PAGING_SIZE].concat(_items.slice(index, index + PAGING_SIZE)));

        this.setState({
          items: itemsCopy
        });
      }, PAGING_DELAY);
    }
  }

  private _onToggleLazyLoad() {
    let { isLazyLoaded } = this.state;

    isLazyLoaded = !isLazyLoaded;

    this.setState({
      isLazyLoaded: isLazyLoaded,
      items: isLazyLoaded ? _items.slice(0, PAGING_SIZE).concat(new Array(ITEMS_COUNT - PAGING_SIZE)) : _items
    });
  }

  private _onToggleResizing() {
    let { items, canResizeColumns, sortedColumnKey, isSortedDescending } = this.state;

    canResizeColumns = !canResizeColumns;

    this.setState({
      canResizeColumns: canResizeColumns,
      columns: this._buildColumns(canResizeColumns, this._onColumnClick, sortedColumnKey, isSortedDescending)
    });
  }

  private _onLayoutChanged(menuItem: IContextualMenuItem) {
    this.setState({
      layoutMode: menuItem.data
    });
  }

  private _onConstrainModeChanged(menuItem: IContextualMenuItem) {
    this.setState({
      constrainMode: menuItem.data
    });
  }

  private _onSelectionChanged(menuItem: IContextualMenuItem) {
    this.setState({
      selectionMode: menuItem.data
    });
  }

  private _getCommandItems() {
    let { layoutMode, constrainMode, selectionMode, canResizeColumns, isLazyLoaded, isHeaderVisible } = this.state;

    return [
      {
        key: 'configure',
        name: 'Configure',
        icon: 'gear',
        items: [
          {
            key: 'resizing',
            name: 'Allow column resizing',
            canCheck: true,
            isChecked: canResizeColumns,
            onClick: this._onToggleResizing
          },
          {
            key: 'headerVisible',
            name: 'Is header visible',
            canCheck: true,
            isChecked: isHeaderVisible,
            onClick: () => this.setState({ isHeaderVisible: !isHeaderVisible })
          },
          {
            key: 'lazyload',
            name: 'Simulate async loading',
            canCheck: true,
            isChecked: isLazyLoaded,
            onClick: this._onToggleLazyLoad
          },
          {
            key: 'dash',
            name: '-'
          },
          {
            key: 'layoutMode',
            name: 'Layout mode',
            items: [
              {
                key: LayoutMode[LayoutMode.fixedColumns],
                name: 'Fixed columns',
                canCheck: true,
                isChecked: layoutMode === LayoutMode.fixedColumns,
                onClick: this._onLayoutChanged,
                data: LayoutMode.fixedColumns
              },
              {
                key: LayoutMode[LayoutMode.justified],
                name: 'Justified columns',
                canCheck: true,
                isChecked: layoutMode === LayoutMode.justified,
                onClick: this._onLayoutChanged,
                data: LayoutMode.justified
              }
            ]
          },
          {
            key: 'selectionMode',
            name: 'Selection mode',
            items: [
              {
                key: SelectionMode[SelectionMode.none],
                name: 'None',
                canCheck: true,
                isChecked: selectionMode === SelectionMode.none,
                onClick: this._onSelectionChanged,
                data: SelectionMode.none

              },
              {
                key: SelectionMode[SelectionMode.single],
                name: 'Single select',
                canCheck: true,
                isChecked: selectionMode === SelectionMode.single,
                onClick: this._onSelectionChanged,
                data: SelectionMode.single
              },
              {
                key: SelectionMode[SelectionMode.multiple],
                name: 'Multi select',
                canCheck: true,
                isChecked: selectionMode === SelectionMode.multiple,
                onClick: this._onSelectionChanged,
                data: SelectionMode.multiple
              },
            ]
          },
          {
            key: 'constrainMode',
            name: 'Constrain mode',
            items: [
              {
                key: ConstrainMode[ConstrainMode.unconstrained],
                name: 'Unconstrained',
                canCheck: true,
                isChecked: constrainMode === ConstrainMode.unconstrained,
                onClick: this._onConstrainModeChanged,
                data: ConstrainMode.unconstrained
              },
              {
                key: ConstrainMode[ConstrainMode.horizontalConstrained],
                name: 'Horizontal constrained',
                canCheck: true,
                isChecked: constrainMode === ConstrainMode.horizontalConstrained,
                onClick: this._onConstrainModeChanged,
                data: ConstrainMode.horizontalConstrained
              }
            ]
          }
        ]
      }
    ];
  }

  private _getContextualMenuProps(column: IColumn, ev: React.MouseEvent): IContextualMenuProps {
    let items = [
      {
        key: 'aToZ',
        name: 'A to Z',
        icon: 'arrowUp2',
        canCheck: true,
        isChecked: column.isSorted && !column.isSortedDescending,
        onClick: () => this._onSortColumn(column.key, false)
      },
      {
        key: 'zToA',
        name: 'Z to A',
        icon: 'arrowDown2',
        canCheck: true,
        isChecked: column.isSorted && column.isSortedDescending,
        onClick: () => this._onSortColumn(column.key, true)
      }
    ];
    return {
      items: items,
      targetElement: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 10,
      isBeakVisible: true,
      onDismiss: this._onContextualMenuDismissed
    };
  }

  private _onItemInvoked(item: any, index: number) {
    console.log('Item invoked', item, index);
  }

  private _onColumnClick(column: IColumn, ev: React.MouseEvent) {
    this.setState({
      contextualMenuProps: this._getContextualMenuProps(column, ev)
    });
  }

  private _onContextualMenuDismissed() {
    this.setState({
      contextualMenuProps: null
    });
  }

  private _onSortColumn(key: string, isSortedDescending: boolean) {
    let sortedItems = this.sort(_items, key, isSortedDescending);

    this.setState({
      items: sortedItems,
      columns: this._buildColumns(true, this._onColumnClick, key, isSortedDescending),
      isSortedDescending: isSortedDescending,
      sortedColumnKey: key
    });
  }

  private sort(items: any[], key: string, isSortedDescending: boolean): any[] {
    return items.sort((a, b) => (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1)
  }

  private _buildColumns(
    canResizeColumns?: boolean,
    onColumnClick?: (column: IColumn, ev: React.MouseEvent) => any,
    sortedColumnKey?: string,
    isSortedDescending?: boolean): IColumn[] {

    return [
      { 
        fieldName: "name",
        key: "name",
        name: "Name",
        onRender: item => <Link>{ item.isFolder ? "Folder: " : "" }{ item.name }</Link>,
        minWidth: 60,
        maxWidth: 400,
        isResizable: canResizeColumns,
        onColumnClick,
        isSorted: sortedColumnKey === "name",
        isSortedDescending,
      },
      {
        fieldName: "lastChanged",
        key: "lastChanged",
        name: "Changed",
        minWidth: 60,
        maxWidth: 180,
        isResizable: canResizeColumns,
        onColumnClick,
        isSorted: sortedColumnKey === "lastChanged",
        isSortedDescending,
      },
      {
        fieldName: "comments",
        key: "comments",
        name: "Comments",
        minWidth: 60,
        maxWidth: 300,
        isCollapsable: true,
        isResizable: canResizeColumns,
        onColumnClick,
        isSorted: sortedColumnKey === "comments",
        isSortedDescending,
      },
    ];
  }

  private createListItems(count: number) {
    const result = [];
    
    for (var index = 0; index < count; index++) {
      result.push(this.createItem(index));
    }

    return result;
  }

  private createItem(index: number) {
    const isFolder = index < 15;
    return {
      name: (isFolder ? "folder" : "file") + index.toString(),
      isFolder,
      lastChanged: new Date("2016/06/" + (1 + index % 30).toString()),
      comments: "PR ######: blah blah blah",
    };
  }
}