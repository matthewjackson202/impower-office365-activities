using System.Activities;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveItemActivity : SharepointDriveActivity
    {
        [Category("Input")]
        [DisplayName("DriveItem ID")]
        [RequiredArgument]
        public InArgument<string> DriveItemID { get; set; }
        internal string driveItemId;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            driveItemId = context.GetValue(DriveItemID);
        }
    }
}
