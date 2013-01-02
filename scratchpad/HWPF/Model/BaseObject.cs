using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Collections;

namespace NPOI.HWPF.Model
{
    /// <summary>

    /// BaseObject class is an abstract class for you to derive from.

    /// Every class that will be dirived from this class will support the 

    /// Clone method automaticly.

    /// The class implements the interface ICloneable and there 

    /// for every object that will be derived 

    /// from this object will support the ICloneable interface as well.

    /// </summary>


    public abstract class BaseObject : ICloneable
    {
        /// <summary>

        /// Clone the object, and returning a reference to a cloned object.

        /// </summary>

        /// <returns>Reference to the new cloned 

        /// object.</returns>

        public virtual object Clone()
        {
            //First we create an instance of this specific type.

            object newObject = Activator.CreateInstance(this.GetType());

            //We get the array of fields for the new type instance.

            FieldInfo[] fields = newObject.GetType().GetFields();

            int i = 0;

            foreach (FieldInfo fi in this.GetType().GetFields())
            {
                //We query if the fiels support the ICloneable interface.

                Type ICloneType = fi.FieldType.
                            GetInterface("ICloneable", true);

                if (ICloneType != null)
                {
                    //Getting the ICloneable interface from the object.

                    ICloneable IClone = (ICloneable)fi.GetValue(this);

                    //We use the clone method to set the new value to the field.

                    fields[i].SetValue(newObject, IClone.Clone());
                }
                else
                {
                    // If the field doesn't support the ICloneable 

                    // interface then just set it.

                    fields[i].SetValue(newObject, fi.GetValue(this));
                }

                //Now we check if the object support the 

                //IEnumerable interface, so if it does

                //we need to enumerate all its items and check if 

                //they support the ICloneable interface.

                Type IEnumerableType = fi.FieldType.GetInterface
                                ("IEnumerable", true);
                if (IEnumerableType != null)
                {
                    //Get the IEnumerable interface from the field.

                    IEnumerable IEnum = (IEnumerable)fi.GetValue(this);

                    //This version support the IList and the 

                    //IDictionary interfaces to iterate on collections.

                    Type IListType = fields[i].FieldType.GetInterface
                                        ("IList", true);
                    Type IDicType = fields[i].FieldType.GetInterface
                                        ("IDictionary", true);

                    int j = 0;
                    if (IListType != null)
                    {
                        //Getting the IList interface.

                        IList list = (IList)fields[i].GetValue(newObject);

                        foreach (object obj in IEnum)
                        {
                            //Checking to see if the current item 

                            //support the ICloneable interface.

                            ICloneType = obj.GetType().
                                GetInterface("ICloneable", true);

                            if (ICloneType != null)
                            {
                                //If it does support the ICloneable interface, 

                                //we use it to set the clone of

                                //the object in the list.

                                ICloneable clone = (ICloneable)obj;

                                list[j] = clone.Clone();
                            }

                            //NOTE: If the item in the list is not 

                            //support the ICloneable interface then in the 

                            //cloned list this item will be the same 

                            //item as in the original list

                            //(as long as this type is a reference type).


                            j++;
                        }
                    }
                    else if (IDicType != null)
                    {
                        //Getting the dictionary interface.

                        IDictionary dic = (IDictionary)fields[i].
                                            GetValue(newObject);
                        j = 0;

                        foreach (DictionaryEntry de in IEnum)
                        {
                            //Checking to see if the item 

                            //support the ICloneable interface.

                            ICloneType = de.Value.GetType().
                                GetInterface("ICloneable", true);

                            if (ICloneType != null)
                            {
                                ICloneable clone = (ICloneable)de.Value;

                                dic[de.Key] = clone.Clone();
                            }
                            j++;
                        }
                    }
                }
                i++;
            }
            return newObject;
        }
    }

}