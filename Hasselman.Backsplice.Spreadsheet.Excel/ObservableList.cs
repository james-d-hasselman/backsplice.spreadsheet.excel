// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

#if DEBUG
[assembly: InternalsVisibleTo("UnitTests")]
#endif

namespace Hasselman.Backsplice.Spreadsheet.Excel
{


    public class ObservableList<T> : IList<T>
    {
        public int Count => values.Count;

        public bool IsReadOnly => false;

        [System.Runtime.CompilerServices.IndexerName("Cells")]
        public T this[int index]
        {
            get => values[index];
            set
            {
                var oldValue = values[index];
                values[index] = value;
                if (ItemUpdated != null)
                {
                    ItemUpdated(this, index, value);
                }
            }
        }

        public delegate void AddEventHandler(object? sender, T item);
        public delegate void ClearEventHandler(object? sender);
        public delegate void InsertEventHandler(object? sender, int index, T item);
        public delegate void ItemUpdatedEventHandler(object? sender, int index, T item);
        public delegate void RemoveEventHandler(object? sender, T item);
        public delegate void RemoveAtEventHandler(object? sender, int index);

        public event AddEventHandler? ItemAdded;
        public event ClearEventHandler? ListCleared;
        public event InsertEventHandler? ItemInserted;
        public event ItemUpdatedEventHandler? ItemUpdated;
        public event RemoveEventHandler? ItemRemoved;
        public event RemoveAtEventHandler? ItemRemovedAt;

        private List<T> values;

        public ObservableList()
        {
            values = new List<T>();
        }

        public ObservableList(IEnumerable<T> collection)
        {
            values = new List<T>(collection);
        }

        public int IndexOf(T item)
        {
            return values.IndexOf(item);
        }

        public void Insert(int index, T item)
        {
            values.Insert(index, item);
            if (ItemInserted != null)
            {
                ItemInserted(this, index, item);
            }
        }

        public void RemoveAt(int index)
        {
            values.RemoveAt(index);
            if(ItemRemovedAt != null)
            {
                ItemRemovedAt(this, index);
            }
        }

        public void Add(T item)
        {
            values.Add(item);
            if(ItemAdded != null)
            {
                ItemAdded(this, item);
            }
        }

        public void Clear()
        {
            values.Clear();
            if(ListCleared != null)
            {
                ListCleared(this);
            }
        }

        public bool Contains(T item)
        {
            return values.Contains(item);
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            values.CopyTo(array, arrayIndex);
        }

        public bool Remove(T item)
        {
            var wasRemoved = values.Remove(item);
            if(wasRemoved && ItemRemoved != null)
            {
                ItemRemoved(this, item);
            }
            return wasRemoved;
        }

        public IEnumerator<T> GetEnumerator()
        {
            return values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return values.GetEnumerator();
        }
    }
}
