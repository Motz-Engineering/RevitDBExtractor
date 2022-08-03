﻿using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

// (c) Revit Database Explorer https://github.com/NeVeSpl/RevitDBExplorer/blob/main/license.md

namespace RevitDBExplorer.Domain.DataModel.MemberAccessors
{
    internal class Schema_EraseSchemaAndAllEntities : MemberAccessorTypedWithWrite<Schema>
    {
        public override ReadResult Read(SnoopableContext context, Schema @object)
        {
            return new ReadResult()
            {
                CanBeSnooped = false,
                Label = $"Erase",
                AccessorName = nameof(Schema_EraseSchemaAndAllEntities)
            };
        }

        public override bool CanBeWritten(SnoopableContext context, Schema schema)
        {
            var result = schema.WriteAccessGranted()&& schema.ReadAccessGranted();
            return result;
        }

        public override Task Write(SnoopableContext context, Schema schema)
        {            
            return context.Execute(x =>
            {   
                var elements = new FilteredElementCollector(context.Document).WherePasses(new ExtensibleStorageFilter(schema.GUID)).ToElements();
                foreach (var element in elements)
                {
                    element.DeleteEntity(schema);
                }
                x.EraseSchemaAndAllEntities(schema); // does not work usually
            }, nameof(Schema_EraseSchemaAndAllEntities));
        }
    }
}
