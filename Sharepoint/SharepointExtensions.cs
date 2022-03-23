using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using Impower.Office365.Sharepoint.Models;

namespace Impower.Office365.Sharepoint
{
    public static class SharepointExtensions
    {
        //TODO: Make this better?
        public static string GetSharepointHostNameFromUrl(string url)
        {
            return url.Replace("http://", String.Empty).Replace("https://", String.Empty).Split('/')[0];
        }
        //TODO: Make this better?
        public static string GetSharepointSitePathFromUrl(string url)
        {
            int index = url.IndexOf("/sites/");
            if (index < 0)
            {
                throw new Exception("Could not find site path from URL");
            }
            return url.Substring(index).TrimEnd('/');
        }
        public static string GetEncodedSharingUrl(string url)
        {
            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(url));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            return encodedUrl;
        }
        public static async Task<Permission> ShareDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            string driveItemId,
            string siteId,
            string driveId,
            LinkType type
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }
            return await drive.Items[driveItemId].CreateLink(type.ToString(), "organization").Request().PostAsync(token);
        }
        public static async Task<DriveItem> GetDriveItemFromSharingUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string sharingUrl
        )
        {
            var encodedUrl = GetEncodedSharingUrl(sharingUrl);
            return await client.Shares[encodedUrl].DriveItem.Request().GetAsync(token);

        }
        public static async Task<Site> GetSiteFromSharingUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string sharingUrl
        )
        {
            var encodedUrl = GetEncodedSharingUrl(sharingUrl);
            return await client.Shares[encodedUrl].Site.Request().GetAsync(token);

        }
        public static async Task<ListItem> GetListItemFromSharingUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string sharingUrl
        )
        {
            var encodedUrl = GetEncodedSharingUrl(sharingUrl);
            return await client.Shares[encodedUrl].ListItem.Request().GetAsync(token);

        }
        public static async Task<Site> GetSharepointSite(
            this GraphServiceClient client,
            CancellationToken token,
            string webUrl
        )
        {
            var hostName = GetSharepointHostNameFromUrl(webUrl);
            var sitePath = GetSharepointSitePathFromUrl(webUrl);
            try {
                return await client.Sites.GetByPath(sitePath, hostName).Request().GetAsync(token);
            }catch(Exception e)
            {
                throw new Exception($"Could not find a site for '{webUrl}'", e);
            }
        }
        public static async Task<bool> GetPermissions(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId
        )
        {
            IDriveRequestBuilder request;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                request = client.Sites[siteId].Drive;
            }
            else
            {
                request = client.Sites[siteId].Drives[driveId];
            }
            var permission = await request.Items[itemId].Permissions.Request().GetAsync(token);
            return permission.ToList();
        }
        public static async Task<Drive> GetSharepointDrive(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveName
        )
        {
            var site = await client.Sites[siteId].Request().GetAsync(token);
            var allDrives = await client.Sites[siteId].Drives.Request().GetAsync(token);
            var matchingDrives = allDrives.Where(drive => drive.Name == driveName);
            if (matchingDrives.Any())
            {
                return matchingDrives.First();
            }
            else
            {
                throw new Exception("Cannot find matching drive.");
            }
        }
        public static async Task<DriveItem> GetSharepointDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }
            return await drive.Items[itemId].Request().Expand(item => item.ListItem).GetAsync(token);
        }
        public static async Task<List> GetSharepointList(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId
        )
        {
            return await client.Sites[siteId].Lists[listId].Request().GetAsync(token);

        }
        public static async Task<FieldValueSet> UpdateSharepointDriveItemFields(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId,
            FieldValueSet fieldValueSet
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }

            return await drive.Items[itemId].ListItem.Fields.Request().UpdateAsync(fieldValueSet,token);

        }
        public static async Task<ListItem> GetSharepointListItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId,
            string itemId
        )
        {
            return await client.Sites[siteId].Lists[listId].Items[itemId].Request().GetAsync(token);
        }
        public static async Task<List<DriveItem>> GetSharepointDriveItemsByPath(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string path
        )
        {
            IDriveItemChildrenCollectionRequest request;
            IDriveRequestBuilder drive;
            string folder = Path.GetDirectoryName(path);
            string filename = Path.GetFileName(path);

            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }
            
            if (String.IsNullOrWhiteSpace(folder))
            {
                request = drive.Root.Children.Request();
            }
            else
            {
                request = drive.Root.ItemWithPath(folder).Children.Request();
            }

            var items = await request.GetAsync(token);
            if (String.IsNullOrWhiteSpace(filename))
            {
                return items.ToList();
            }
            else
            {
                return items.Where(item => item.Name == filename).ToList();
            }

        }
    }
}
