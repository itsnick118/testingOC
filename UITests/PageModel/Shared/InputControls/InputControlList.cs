using System.Collections;
using System.Collections.Generic;

namespace UITests.PageModel.Shared.InputControls
{
    public class InputControlList : IEnumerable<InputControl>
    {
        private readonly IDictionary<string, InputControl> _controls;

        public InputControlList()
        {
            _controls = new Dictionary<string, InputControl>();
        }

        public static InputControlList FromEnumerable(IEnumerable<InputControl> controls)
        {
            var controlList = new InputControlList();
            foreach (var control in controls)
            {
                controlList.Add(control);
            }

            return controlList;
        }

        public IDictionary<string, InputControl> ToDictionary()
        {
            return _controls;
        }

        public IEnumerator<InputControl> GetEnumerator()
        {
            return _controls.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Count => _controls.Count;

        public bool ContainsKey(string key)
        {
            return _controls.ContainsKey(key);
        }

        public void Add(InputControl item)
        {
            _controls.Add(item.Name, item);
        }

        public bool Remove(string key)
        {
            return _controls.Remove(key);
        }

        public bool TryGetValue(string key, out InputControl value)
        {
            return _controls.TryGetValue(key, out value);
        }

        public InputControl this[string key]
        {
            get => _controls[key];
            set => _controls[key] = value;
        }

        public ICollection<string> Keys => _controls.Keys;

        public ICollection<InputControl> Values => _controls.Values;
    }
}