# frozen_string_literal: true

include BrowserConfigHelper

describe BrowseEverything::Driver::Box do
  subject { provider }

  let(:browser) { BrowseEverything::Browser.new(url_options) }
  let(:provider) { browser.providers['box'] }
  let(:auth_params) do
    {
      'code' => 'CODE',
      'state' => 'box'
    }
  end
  let(:token) do
    {
      'token' => 'TOKEN',
      'refresh_token' => 'REFRESH_TOKEN'
    }
  end
  let(:oauth_response_body) { '{"access_token":"TOKEN","expires_in":3762,"restricted_to":[],"refresh_token":"REFRESH_TOKEN","token_type":"bearer"}' }

  before do
    stub_configuration

    stub_request(
      :post,
      'https://www.box.com/api/oauth2/token'
    ).to_return(
      body: oauth_response_body,
      status: 200,
      headers: {
        'Content-Type' => 'application/json'
      }
    )
  end

  after do
    unstub_configuration
  end

  its(:name) { is_expected.to eq('Box') }
  its(:key)  { is_expected.to eq('box') }
  its(:icon) { is_expected.to be_a(String) }

  describe '#validate_config' do
    it 'raises and error with an incomplete configuration' do
      expect { described_class.new({}) }.to raise_error(BrowseEverything::InitializationError)
    end
  end

  describe '#auth_link' do
    subject { provider.auth_link }

    it { is_expected.to start_with('https://www.box.com/api/oauth2/authorize') }
    it { is_expected.to include('browse%2Fconnect') }
    it { is_expected.to include('response_type') }
  end

  describe '#authorized?' do
    subject { provider.authorized? }

    context 'when the access token is not registered' do
      it { is_expected.to be(false) }
    end

    context 'when the access tokens are registered and not expired' do
      before { provider.token = token.merge('expires_at' => Time.now.to_i + 360) }

      it { is_expected.to be(true) }
    end

    context 'when the access tokens are registered but no expiration time' do
      before { provider.token = token }

      it { is_expected.to be(false) }
    end

    context 'when the access tokens are registered but expired' do
      before { provider.token = token.merge('expires_at' => Time.now.to_i - 360) }

      it { is_expected.to be(false) }
    end
  end

  describe '#connect' do
    before do
      allow(provider).to receive(:register_access_token)
    end

    it 'registers new tokens' do
      provider.connect(auth_params, 'data', nil)
      expect(provider).to have_received(:register_access_token).with(kind_of(OAuth2::AccessToken))
    end
  end

  describe '#contents' do
    let(:folders_response_body) do
      '{"type":"folder","id":"0","sequence_id":null,"etag":null,"name":"All Files","created_at":null,"modified_at":null,"description":"","size":28747877,"path_collection":{"total_count":0,"entries":[]},"created_by":{"type":"user","id":"","name":"","login":""},"modified_by":{"type":"user","id":"225383863","name":"ADAM GARNER WEAD","login":"agw13@psu.edu"},"trashed_at":null,"purged_at":null,"content_created_at":null,"content_modified_at":null,"owned_by":{"type":"user","id":"225383863","name":"ADAM GARNER WEAD","login":"agw13@psu.edu"},"shared_link":null,"folder_upload_email":null,"parent":null,"item_status":"active","item_collection":{"total_count":13,"entries":[{"type":"folder","id":"20375782799","sequence_id":"0","etag":"0","name":"A very looooooooooooong box folder, why so loooooooong Lets make it even longer to show how far it gets sent to the side"},{"type":"folder","id":"2571160559","sequence_id":"0","etag":"0","name":"Apps Team - Shared"},{"type":"folder","id":"20194542723","sequence_id":"0","etag":"0","name":"DSRD - W Pattee 3"},{"type":"folder","id":"20284062015","sequence_id":"0","etag":"0","name":"My Box Notes"},{"type":"folder","id":"11305958926","sequence_id":"0","etag":"0","name":"PCDM-Sufia"},{"type":"folder","id":"4227519189","sequence_id":"0","etag":"0","name":"refactor"},{"type":"folder","id":"2459961273","sequence_id":"0","etag":"0","name":"SaS - Development Team"},{"type":"folder","id":"3399219062","sequence_id":"0","etag":"0","name":"Scholarsphere - Migration"},{"type":"folder","id":"1168461187","sequence_id":"0","etag":"0","name":"test"},{"type":"folder","id":"3055812547","sequence_id":"0","etag":"0","name":"UX Artifacts"},{"type":"file","id":"25581309763","file_version":{"type":"file_version","id":"23869158869","sha1":"4604bbe44fdcdd4afef3c666cf582e3773960954"},"sequence_id":"1","etag":"1","sha1":"4604bbe44fdcdd4afef3c666cf582e3773960954","name":"failed.tar.gz"},{"type":"file","id":"25588823531","file_version":{"type":"file_version","id":"23877641673","sha1":"abd18ce0a685a27b464fb05f27af5a84f9ec9be7"},"sequence_id":"1","etag":"1","sha1":"abd18ce0a685a27b464fb05f27af5a84f9ec9be7","name":"scholarsphere_5712md360.xml"},{"type":"file","id":"113711622968","file_version":{"type":"file_version","id":"122136107320","sha1":"da39a3ee5e6b4b0d3255bfef95601890afd80709"},"sequence_id":"0","etag":"0","sha1":"da39a3ee5e6b4b0d3255bfef95601890afd80709","name":"test.txt"}],"offset":0,"limit":100,"order":[{"by":"type","direction":"ASC"},{"by":"name","direction":"ASC"}]}}'
    end
    let(:folders_items_response_body) do
      '{"total_count":13,"entries":[{"type":"folder","id":"20375782799","etag":"0","name":"A very looooooooooooong box folder, why so loooooooong Lets make it even longer to show how far it gets sent to the side","size":0,"created_at":"2017-03-01T04:15:15-08:00"},{"type":"folder","id":"2571160559","etag":"0","name":"Apps Team - Shared","size":1249,"created_at":"2014-10-15T13:00:29-07:00"},{"type":"folder","id":"20194542723","etag":"0","name":"DSRD - W Pattee 3","size":2949416,"created_at":"2017-02-27T08:17:21-08:00"},{"type":"folder","id":"20284062015","etag":"0","name":"My Box Notes","size":0,"created_at":"2017-02-28T08:52:26-08:00"},{"type":"folder","id":"11305958926","etag":"0","name":"PCDM-Sufia","size":650658,"created_at":"2016-09-14T09:14:25-07:00"},{"type":"folder","id":"4227519189","etag":"0","name":"refactor","size":8766,"created_at":"2015-08-14T07:53:56-07:00"},{"type":"folder","id":"2459961273","etag":"0","name":"SaS - Development Team","size":152720753,"created_at":"2014-09-17T13:39:31-07:00"},{"type":"folder","id":"3399219062","etag":"0","name":"Scholarsphere - Migration","size":270984,"created_at":"2015-04-07T13:17:51-07:00"},{"type":"folder","id":"1168461187","etag":"0","name":"test","size":20625445557,"created_at":"2013-09-19T12:57:59-07:00"},{"type":"folder","id":"3055812547","etag":"0","name":"UX Artifacts","size":3801994,"created_at":"2015-02-04T08:21:16-08:00"},{"type":"file","id":"25581309763","etag":"1","name":"failed.tar.gz","size":28650839,"created_at":"2015-01-29T05:18:43-08:00"},{"type":"file","id":"25588823531","etag":"1","name":"scholarsphere_5712md360.xml","size":97038,"created_at":"2015-01-29T08:38:44-08:00"},{"type":"file","id":"113711622968","etag":"0","name":"test.txt","size":0,"created_at":"2016-12-20T07:50:30-08:00"}],"offset":0,"limit":1000,"order":[{"by":"type","direction":"ASC"},{"by":"name","direction":"ASC"}]}'
    end
    before do
      provider.token = token

      stub_request(
        :get, "https://api.box.com/2.0/folders/0"
      ).to_return(
        body: folders_response_body,
        status: 200,
        headers: {
          'Content-Type' => 'application/json'
        }
      )
      stub_request(
        :get, "https://api.box.com/2.0/folders/0/items?fields=name,size,created_at&limit=99999&offset=0"
      ).to_return(
        body: folders_items_response_body,
        status: 200,
        headers: {
          'Content-Type' => 'application/json'
        }
      )
    end

    # HERE
    context 'with files and folders in the root directory' do
      let(:root_directory) { provider.contents('') }
      let(:long_file)      { root_directory[0] }
      let(:sas_directory)  { root_directory[6] }
      let(:tar_file)       { root_directory[10] }

      describe 'the first item' do
        subject { long_file }

        its(:name)     { is_expected.to start_with('A very looooooooooooong box folder') }
        its(:location) { is_expected.to eq('box:20375782799') }
        it             { is_expected.to be_container }
      end

      describe 'the SaS - Development Team directory' do
        subject { sas_directory }

        its(:name)     { is_expected.to eq('SaS - Development Team') }
        its(:location) { is_expected.to eq('box:2459961273') }
        its(:id)       { is_expected.to eq('2459961273') }
        it             { is_expected.to be_container }
      end

      describe 'a file' do
        subject { tar_file }

        its(:name)     { is_expected.to eq('failed.tar.gz') }
        its(:size)     { is_expected.to eq(28_650_839) }
        its(:location) { is_expected.to eq('box:25581309763') }
        its(:type)     { is_expected.to eq('application/x-gzip') }
        its(:id)       { is_expected.to eq('25581309763') }
        it             { is_expected.not_to be_container }
      end
    end

    context 'with files and folders in the SaS - Development Team directory' do
      let(:folder_id) { '2459961273' }
      let(:sas_directory)    { provider.contents(folder_id) }
      let(:apps_dir)         { sas_directory[0] }
      let(:equipment)        { sas_directory[11] }
      let(:folders_response_body) do
        '{"type":"folder","id":"2459961273","sequence_id":"0","etag":"0","name":"SaS - Development Team","created_at":"2014-09-17T13:39:31-07:00","modified_at":"2017-03-17T08:41:40-07:00","description":"","size":152720753,"path_collection":{"total_count":1,"entries":[{"type":"folder","id":"0","sequence_id":null,"etag":null,"name":"All Files"}]},"created_by":{"type":"user","id":"191882215","name":"Daniel Coughlin","login":"dmc186@psu.edu"},"modified_by":{"type":"user","id":"190902775","name":"CAROLYN A COLE","login":"cam156@psu.edu"},"trashed_at":null,"purged_at":null,"content_created_at":"2014-09-17T13:39:31-07:00","content_modified_at":"2017-03-17T08:41:40-07:00","owned_by":{"type":"user","id":"191882215","name":"Daniel Coughlin","login":"dmc186@psu.edu"},"shared_link":{"url":"https:\/\/psu.box.com\/s\/8hb3e06nthwld39ehlg5","download_url":null,"vanity_url":null,"effective_access":"collaborators","is_password_enabled":false,"unshared_at":null,"download_count":0,"preview_count":0,"access":"collaborators","permissions":{"can_download":true,"can_preview":true}},"folder_upload_email":null,"parent":null,"item_status":"active","item_collection":{"total_count":17,"entries":[{"type":"folder","id":"2459974427","sequence_id":"0","etag":"0","name":"Apps&Int"},{"type":"folder","id":"11217040834","sequence_id":"0","etag":"0","name":"Credentials"},{"type":"folder","id":"5662231069","sequence_id":"0","etag":"0","name":"DevOps Presentation"},{"type":"folder","id":"8648394509","sequence_id":"0","etag":"0","name":"old"},{"type":"folder","id":"12445061821","sequence_id":"0","etag":"0","name":"Ops Plan"},{"type":"file","id":"136917954839","file_version":{"type":"file_version","id":"158227338595","sha1":"2e2124c0d6589dd05d03a854b9c5c3a68bbc4501"},"sequence_id":"17","etag":"17","sha1":"2e2124c0d6589dd05d03a854b9c5c3a68bbc4501","name":"2-17 Upcoming ScholarSphere changes.boxnote"},{"type":"file","id":"75228871930","file_version":{"type":"file_version","id":"79668751214","sha1":"43cd4ac728cef7767dc132c26b181133d867d78f"},"sequence_id":"1","etag":"1","sha1":"43cd4ac728cef7767dc132c26b181133d867d78f","name":"99bottles (1).epub"},{"type":"file","id":"72329914177","file_version":{"type":"file_version","id":"76429987577","sha1":"a3310fdad0e6c9654b19d88b35318b8cdc246379"},"sequence_id":"2","etag":"2","sha1":"a3310fdad0e6c9654b19d88b35318b8cdc246379","name":"Development Team Projects and Milestones (not downloaded).xlsx"},{"type":"file","id":"74507436394","file_version":{"type":"file_version","id":"78854145518","sha1":"589ab64efb1ba49537074fd2c56e1570a24a47b5"},"sequence_id":"7","etag":"7","sha1":"589ab64efb1ba49537074fd2c56e1570a24a47b5","name":"Development Team Projects and Milestones - Editable.xlsx"},{"type":"file","id":"72324078305","file_version":{"type":"file_version","id":"76423451525","sha1":"42c0a5a34717cd17e315653643b61d23e94b7cd7"},"sequence_id":"3","etag":"3","sha1":"42c0a5a34717cd17e315653643b61d23e94b7cd7","name":"Development Team Projects and Milestones.xlsx"},{"type":"file","id":"106651625938","file_version":{"type":"file_version","id":"114805197783","sha1":"0c1c90a0186bd3934ce3877e092a228dc1568e93"},"sequence_id":"2","etag":"2","sha1":"0c1c90a0186bd3934ce3877e092a228dc1568e93","name":"Digital Scholarship and Repository Development (DSRD) Org Chart 2016.pdf"},{"type":"file","id":"76960974625","file_version":{"type":"file_version","id":"106919307872","sha1":"74608bfb532feff774502c6132eafe9acb1e5217"},"sequence_id":"162","etag":"162","sha1":"74608bfb532feff774502c6132eafe9acb1e5217","name":"Equipment.boxnote"},{"type":"file","id":"68362464289","file_version":{"type":"file_version","id":"101137863811","sha1":"40a0a3f98926874f1ee62fd522621a4408e0058c"},"sequence_id":"8","etag":"8","sha1":"40a0a3f98926874f1ee62fd522621a4408e0058c","name":"migrationDesign.pptx"},{"type":"file","id":"99028795390","file_version":{"type":"file_version","id":"107183805224","sha1":"7044bd7f0892f915b77b914430c2915161f942d9"},"sequence_id":"37","etag":"37","sha1":"7044bd7f0892f915b77b914430c2915161f942d9","name":"Onboarding.boxnote"},{"type":"file","id":"134865535934","file_version":{"type":"file_version","id":"143971901576","sha1":"32cc33b2cd5cc46f0a3fb924b4779cc387a81018"},"sequence_id":"1","etag":"1","sha1":"32cc33b2cd5cc46f0a3fb924b4779cc387a81018","name":"passenger-test.boxnote"},{"type":"file","id":"92102699797","file_version":{"type":"file_version","id":"98754405909","sha1":"053b762f6d5752c261ef7ff6a3d776853cdb2ca3"},"sequence_id":"1","etag":"1","sha1":"053b762f6d5752c261ef7ff6a3d776853cdb2ca3","name":"sas_development_libraries_mou_2016_08_01.docx"},{"type":"file","id":"22182934895","file_version":{"type":"file_version","id":"50126669941","sha1":"8ca396b5b26ff90d625ea3a49f0cc69ad844476b"},"sequence_id":"41","etag":"41","sha1":"8ca396b5b26ff90d625ea3a49f0cc69ad844476b","name":"TimeBox.xlsx"}],"offset":0,"limit":100,"order":[{"by":"type","direction":"ASC"},{"by":"name","direction":"ASC"}]}}'
      end
      let(:folders_items_response_body) do
        '{"total_count":17,"entries":[{"type":"folder","id":"2459974427","etag":"0","name":"Apps&Int","size":14341346,"created_at":"2014-09-17T13:42:30-07:00"},{"type":"folder","id":"11217040834","etag":"0","name":"Credentials","size":12603,"created_at":"2016-09-07T08:37:23-07:00"},{"type":"folder","id":"5662231069","etag":"0","name":"DevOps Presentation","size":128158633,"created_at":"2015-12-07T05:51:08-08:00"},{"type":"folder","id":"8648394509","etag":"0","name":"old","size":117232,"created_at":"2016-06-30T08:03:46-07:00"},{"type":"folder","id":"12445061821","etag":"0","name":"Ops Plan","size":6181924,"created_at":"2016-11-17T06:52:58-08:00"},{"type":"file","id":"136917954839","etag":"17","name":"2-17 Upcoming ScholarSphere changes.boxnote","size":1768,"created_at":"2017-02-17T09:27:57-08:00"},{"type":"file","id":"75228871930","etag":"1","name":"99bottles (1).epub","size":1013573,"created_at":"2016-07-21T03:53:01-07:00"},{"type":"file","id":"72329914177","etag":"2","name":"Development Team Projects and Milestones (not downloaded).xlsx","size":20812,"created_at":"2016-06-30T09:10:36-07:00"},{"type":"file","id":"74507436394","etag":"7","name":"Development Team Projects and Milestones - Editable.xlsx","size":21001,"created_at":"2016-07-15T11:08:54-07:00"},{"type":"file","id":"72324078305","etag":"3","name":"Development Team Projects and Milestones.xlsx","size":22410,"created_at":"2016-06-30T08:31:56-07:00"},{"type":"file","id":"106651625938","etag":"2","name":"Digital Scholarship and Repository Development (DSRD) Org Chart 2016.pdf","size":17677,"created_at":"2016-11-30T08:27:58-08:00"},{"type":"file","id":"76960974625","etag":"162","name":"Equipment.boxnote","size":10140,"created_at":"2016-08-03T12:22:39-07:00"},{"type":"file","id":"68362464289","etag":"8","name":"migrationDesign.pptx","size":47751,"created_at":"2016-06-03T03:56:54-07:00"},{"type":"file","id":"99028795390","etag":"37","name":"Onboarding.boxnote","size":4465,"created_at":"2016-10-24T11:41:47-07:00"},{"type":"file","id":"134865535934","etag":"1","name":"passenger-test.boxnote","size":4988,"created_at":"2017-02-13T10:47:38-08:00"},{"type":"file","id":"92102699797","etag":"1","name":"sas_development_libraries_mou_2016_08_01.docx","size":134018,"created_at":"2016-08-16T07:32:53-07:00"},{"type":"file","id":"22182934895","etag":"41","name":"TimeBox.xlsx","size":21864,"created_at":"2014-10-23T08:22:16-07:00"}],"offset":0,"limit":1000,"order":[{"by":"type","direction":"ASC"},{"by":"name","direction":"ASC"}]}'
      end

      before do
        stub_request(
          :get, "https://api.box.com/2.0/folders/#{folder_id}"
        ).to_return(
          body: folders_response_body,
          status: 200,
          headers: {
            'Content-Type' => 'application/json'
          }
        )
        stub_request(
          :get, "https://api.box.com/2.0/folders/#{folder_id}/items?fields=name,size,created_at&limit=99999&offset=0"
        ).to_return(
          body: folders_items_response_body,
          status: 200,
          headers: {
            'Content-Type' => 'application/json'
          }
        )
      end

      describe 'the second item' do
        subject { apps_dir }

        its(:name)     { is_expected.to eq('Apps&Int') }
        its(:id)       { is_expected.to eq('2459974427') }
        it             { is_expected.to be_container }
      end

      describe 'a file' do
        subject { equipment }

        its(:name)     { is_expected.to eq('Equipment.boxnote') }
        its(:size)     { is_expected.to eq(10140) }
        its(:location) { is_expected.to eq('box:76960974625') }
        its(:type)     { is_expected.to eq('application/octet-stream') }
        its(:id)       { is_expected.to eq('76960974625') }
        it             { is_expected.not_to be_container }
      end
    end
  end

  describe '#link_for' do
    before do
      provider.token = token

      # Initial request for metadata
      stub_request(
        :get, "https://api.box.com/2.0/files/#{file_id}"
      ).to_return(
        body: files_response_body,
        status: 200,
        headers: {
          'Content-Type' => 'application/json'
        }
      )

      # Request for the content
      stub_request(
        :get, "https://api.box.com/2.0/files/#{file_id}/content"
      ).to_return(
        status: 302,
        headers: {
          'location' => files_response_location
        }
      )
    end
    let(:link) { provider.link_for(file_id) }

    context 'with a file from the root directory' do
      let(:files_response_location) do
        'https://dl.boxcloud.com/d/1/B7Qd7B_iwPHTU4-71W2qYoTbvAaHCPzNsy5WTFHj5XpbmydlF8ud_0n7Ji7zswdU3Kg0patL8EUXCR3cpPw1PZjp-_a1t6MoH-tX3eQeCAR080Hr-yQaEmLQ8dULnnlOhYHeYwPuRp-0gCcXCF5w3O3bE6ZHgML3SCQchPoJQlsfvcYwXZyPRFEVRNUo7qou5X9dbkMCJGmB0CsvqKTXfEs8bqbmRV4hZ5qpJbD0Jer1m4vsqu8h5VkBdMIcgFMn_D9TheZOrmQdg8ExZZVPJ7X8QjFjI707WwIl7CNYkmocCdAJnbYQljFMiQsp0wF0etUqoiskNnaJBS3NagvtvSKaX3TVfiXa87CHDwPiQ3PNJE32d49eKcBK9nLLigoX3SJOkJqMqUSXO_UAxl0bO0EszCNpaWsiQwWiG2jjB5YBfDwQbmfuSTCBPkhDjIH3S3n5cadk4H_8rvgHshYBXNJd2NshgPsYt1XIkE8qj4WznzwayZoi_2k5x4liFvs2F91anj0bPAW2Mzeyz0Pi49mythoH7Rrq_i0sbYDt3VJNyCB0wqs8xUCa38Z6NocTKrKtFHrNSyM2-g5nWsAIwGeG4L4kZ_qIq8nNL49LsntvpFmhogONBozQQCvoEj0WXqhLCxxwT4U-pZ5rCYDt6dpoNlmYv962wXUzg_s21CO96J1o9-_xKl9JxmbWPf0DDJ35gCBFWMhL6BTaLckbBrybx41luHjjDuv72CxfsqsabdyW2yDN3uoBAcXIXD7QUOACPXKHvC98kL9aXL9o9X-zt89vUP61ijizsjtoI9RQc4CrjerfnxC7jS-ER3w2nwkgOEr1wkdEvq4QabcVlyixgrRxKTmzHXa0JGfmpLdEQ821Sgi7QHOpa0l5CSZTQycnWHoMKn4f_xJo5WZLmkmNgkE2DRsok0xGVUK_lNKjrt_N9mPVAdeeXLjSyoSol2_ugU79ELQSHJKKyTFAEJOscUSseg5MtWrIOLKQY6NlAaE0Ckn7c_LX4MYaSj83F-sZouvwSRNVbTklMNkC6j2fy4kr8_T3eb_9-Df94B1867kSJzt7TjXSAxz_p9PbRuj7eY1NW15zjxtOPQpk/download'
      end
      let(:files_response_body) do
        '{"type":"file","id":"25581309763","file_version":{"type":"file_version","id":"23869158869","sha1":"4604bbe44fdcdd4afef3c666cf582e3773960954"},"sequence_id":"1","etag":"1","sha1":"4604bbe44fdcdd4afef3c666cf582e3773960954","name":"failed.tar.gz","description":"","size":28650839,"path_collection":{"total_count":1,"entries":[{"type":"folder","id":"0","sequence_id":null,"etag":null,"name":"All Files"}]},"created_at":"2015-01-29T05:18:43-08:00","modified_at":"2015-01-29T05:18:43-08:00","trashed_at":null,"purged_at":null,"content_created_at":"2015-01-17T05:59:48-08:00","content_modified_at":"2015-01-17T06:00:03-08:00","created_by":{"type":"user","id":"225383863","name":"ADAM GARNER WEAD","login":"agw13@psu.edu"},"modified_by":{"type":"user","id":"225383863","name":"ADAM GARNER WEAD","login":"agw13@psu.edu"},"owned_by":{"type":"user","id":"225383863","name":"ADAM GARNER WEAD","login":"agw13@psu.edu"},"shared_link":{"url":"https:\/\/psu.box.com\/s\/w89hg6kf73q8fn1ww51cj92ohb6piyet","download_url":"https:\/\/psu.box.com\/shared\/static\/w89hg6kf73q8fn1ww51cj92ohb6piyet.gz","vanity_url":null,"effective_access":"open","is_password_enabled":false,"unshared_at":null,"download_count":1,"preview_count":0,"access":"open","permissions":{"can_download":true,"can_preview":true}},"parent":{"type":"folder","id":"0","sequence_id":null,"etag":null,"name":"All Files"},"item_status":"active"}'
      end
      let(:file_id) { '25581309763' }

      specify { expect(link[0]).to start_with('https://dl.boxcloud.com/d/1') }
      specify { expect(link[1]).to have_key(:expires) }
    end

    context 'with a file from the SaS - Development Team directory' do
      let(:files_response_location) do
        'https://dl.boxcloud.com/d/1/lCA2dDrIV1QAvpQLU7_mkJ0bt2Soa1dLRT6FOzmF37EfJgjRmCO-rZ0VFyKCtoHgPzRoCIHgx-IWCoV8cvtfTX4Yw--RmBLBV9f4JsK6i3LMRKQzgDMxsMu97RSmuMagV-GayR8uO6NGFBtoX81yujebVY-JRB7cxhPU0fbCxAzAnv1711_IUs6YXwhWc-rNFHrNjPsUnJLqw0soQPZqF3Q5irxJLu7traVIeoeuGSfhw-7G-qfqu4CIFonC4ktwwh8jMgN9KW712kRg61moLG6Aa6FTbdCCSt_jCwvQFVGqCz_VivfzG2_BHBqh4IhIB76DmF1ISM-jD9VToUwYfscg6BgC2fqGj6OycsuYr3v3EE0gHlPw_X2rJnq9J2M9plnnhqBvXEwoUXDqyGPvSqWBeQvrrTYfvqj7fG48tUr7RKOytc_K0B4aw9hgGUk3EaifSVWzYvXzoXBp27HFHPsssvIK8fXoSU2J3HCEsGhIDFzu3cvXyPZJV_Co1_LvivWM0AGjXhyoC8PNO5qmtrZFgCob3KlZ8BEI1yJfX_0K8wpocUJP95eBNutBhWa2DEvCK6R5OG-Z0XwrJu8XFRwPcIl3_7lSsIhsQlYA7MoJer7hhwVGAl3kPP-bFsw1et66UM6KHG2FdN2xKP10DYFqR_JrShPP8DQ-ik9-lKej_8LTfiHdf0YbkdMNKzBYlXZsiQNwd3VXdRo1Z4Uhza_XtTDcfH0uxRQyYRecHr-5lEn5iUfV1199GrTsLkcOb9ONNSTFPKW15XCLpgJm6ylSpG2JSZ6zWsGL5hUzdvMytwBCMpIj84UZuwo2RTH1ZDUJYjmHRX11rDnwe2zKMz4woYIA8RjhnYEj6d13BeBz9Fy4GgR90jmM5xHIv9fiAG_b99fpDuYvT-tccN4pdrxPf0P-kCjYAHBt7kY7ToaMWcnOf1nGdLP0oaCGHhAzN0iurjRTLRi2S4N1hsWxjoMuQqayEq15QhRujy2lTZX_XyBgrU2ZodkNsQUlaWcJ6aFT8bh48DtyQvt_xfm4jqmStageFKbydyoB8BpP6oYFAMeFlkS4A6LnXJCzpO8Q../download'
      end
      let(:files_response_body) do
        '{"type":"file","id":"76960974625","file_version":{"type":"file_version","id":"106919307872","sha1":"74608bfb532feff774502c6132eafe9acb1e5217"},"sequence_id":"162","etag":"162","sha1":"74608bfb532feff774502c6132eafe9acb1e5217","name":"Equipment.boxnote","description":"Listing equipment we currently have","size":10140,"path_collection":{"total_count":2,"entries":[{"type":"folder","id":"0","sequence_id":null,"etag":null,"name":"All Files"},{"type":"folder","id":"2459961273","sequence_id":"0","etag":"0","name":"SaS - Development Team"}]},"created_at":"2016-08-03T12:22:39-07:00","modified_at":"2016-10-26T12:40:09-07:00","trashed_at":null,"purged_at":null,"content_created_at":"2016-08-03T12:22:39-07:00","content_modified_at":"2016-10-26T12:40:09-07:00","created_by":{"type":"user","id":"212771336","name":"NICOLE M. GAMPE","login":"nmg110@psu.edu"},"modified_by":{"type":"user","id":"208208274","name":"JUSTIN R PATTERSON","login":"jrp22@psu.edu"},"owned_by":{"type":"user","id":"191882215","name":"Daniel Coughlin","login":"dmc186@psu.edu"},"shared_link":{"url":"https:\/\/psu.box.com\/s\/uolieeszpwglsv0ipwij9ws36rrepvba","download_url":"https:\/\/psu.box.com\/shared\/static\/uolieeszpwglsv0ipwij9ws36rrepvba.boxnote","vanity_url":null,"effective_access":"open","is_password_enabled":false,"unshared_at":null,"download_count":24,"preview_count":0,"access":"open","permissions":{"can_download":true,"can_preview":true}},"parent":{"type":"folder","id":"2459961273","sequence_id":"0","etag":"0","name":"SaS - Development Team"},"item_status":"active"}'
      end
      let(:file_id) { '76960974625' }

      specify { expect(link[0]).to start_with('https://dl.boxcloud.com/d/1') }
      specify { expect(link[1]).to have_key(:expires) }
    end
  end
end
