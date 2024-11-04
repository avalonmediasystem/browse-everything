# frozen_string_literal: true

require_relative 'authentication_factory'

module BrowseEverything
  module Driver
    # Driver for accessing the MS-Graph API (https://learn.microsoft.com/en-us/graph/overview)
    class Sharepoint < Base

      class << self
        attr_accessor :authentication_klass

        def default_authentication_klass
          BrowseEverything::Auth::Sharepoint::Session
        end
      end

      # Constructor
      # @param config_values [Hash] configuration for the driver
      def initialize(config_values)
        self.class.authentication_klass ||= self.class.default_authentication_klass
        super(config_values)
      end

      def icon
        'cloud'
      end

      # Validates the configuration for the Sharepoint provider
      def validate_config
        raise InitializationError, 'Sharepoint driver requires a :client_id argument' unless config[:client_id]
        raise InitializationError, 'Sharepoint driver requires a :client_secret argument' unless config[:client_secret]
        raise InitializationError, 'Sharepoint driver requires a :tenant_id argument' unless config[:tenant_id]
        raise InitializationError, 'Sharepoint driver requires a :redirect_uri argument' unless config[:redirect_uri]
        raise InitializationError, 'Sharepoint driver requires a :scope argument' unless config[:scope]
      end

      # Retrieves the file entry objects for a given path to MS-graph drive resource
      # @param [String] id of the file or folder
      # @return [Array<BrowseEverything::FileEntry>]
      def contents(id = '')
        token_refresh if authorized?

        folder = []
        if id.empty?
          folder << sites
          folder << drives
        else
          folder << items_by_id(id)
        end

        values = []

        folder.flatten.each do |f|
          values << directory_entry(f)
        end
        @entries = values.compact

        @sorter.call(@entries)
      end

      # @return [String]
      # Authorization url that is used to request the initial access code from Sharepoint/Onedrive/365/etc
      def auth_link(*_args)
        Addressable::URI.parse("https://login.microsoftonline.com/#{config[:tenant_id]}/oauth2/v2.0/authorize?#{auth_query_string}")
      end

      # @return [Boolean]
      def authorized?
        @token.present?
      end

      def authorize!
        return if @code.blank?
        register_access_token(sharepoint_session.get_access_token(@code))
        @code = nil
        @token
      end

      def connect(params, _data, _url_options)
        @code = params[:code]
        authorize!
      end

      # @param [String] id of the file on MS graph drive
      # @return [Array<String, Hash>]
      def link_for(id)
        file = items_by_id(id)
        extras = {file_name: file['name'], file_size: file['size'].to_i}
        [download_url(file), extras]
      end

      private

      def auth_query_string
        query = []
        # keep_if deletes from and returns self, so dup config to not overwrite original
        base = config.dup.keep_if { |k,v| ['client_id', 'scope', 'redirect_uri'].include?(k) }
        base.each do |k,v|
          query += ["#{k}=#{v}"]
        end
        query += ["response_type=code"]

        query.join('&')
      end

      def session
        AuthenticationFactory.new(
          self.class.authentication_klass,
          client_id: config[:client_id],
          client_secret: config[:client_secret],
          tenant_id: config[:tenant_id],
          scope: config[:scope],
          redirect_uri: config[:redirect_uri],
          code: @code.presence,
          access_token: @token.presence
        )
      end

      def authenticate
        session.authenticate
      end

      def sharepoint_session
        @sharepoint_session ||= authenticate
      end

      def token_refresh
        return @token unless token_expired?

        register_access_token(sharepoint_session.refresh_token)
      end

      # If there is an active session, {@token} will be set by {BrowseEverythingController} using data stored in the
      # session. 
      #
      # @param [OAuth2::AccessToken] access_token
      def register_access_token(access_token)
        @token = {
                   'token' => access_token.token,
                   'expires_in' => access_token.expires_in,
                   'expires_at' => access_token.expires_at,
                   'refresh_token' => access_token.refresh_token
                 }
      end

      def sharepoint_token
        return unless @token
        @token.fetch('token', nil)
      end

      def expiration_time
        return unless @token
        @token.fetch('expires_at', nil).to_i
      end

      def token_expired?
        return true if expiration_time.nil?
        Time.now.to_i > expiration_time
      end

      def sharepoint_request(sharepoint_uri)
        @auth = "Bearer " + sharepoint_token

        uri = URI.parse(sharepoint_uri)
        http = Net::HTTP.new(uri.host, uri.port)
        http.use_ssl = true if uri.scheme == 'https'

        response = http.start do
          request = Net::HTTP::Get.new(uri.request_uri,{'Authorization' => @auth})
          http.request(request)
        end
        JSON.parse(response.body)
      end

      # Constructs a BrowseEverything::FileEntry object for a Sharepoint file
      # resource
      # @param file [String] ID to the file resource
      # @return [BrowseEverything::File]
      def directory_entry(file)
        BrowseEverything::FileEntry.new(make_path(file), 
                                        [key, make_path(file)].join(':'), 
                                        file['displayName'] ? file['displayName'] : file['name'], 
                                        file['size'] ? file['size'] : nil, 
                                        Date.parse(file['lastModifiedDateTime']),
                                        folder?(file))
      end

      # Derives a path from item (file or folder or drive) metadata 
      # that can be used in subsequent items_by_id calls
      def make_path(file)
        if file['parentReference'].present? 
          folder?(file) ? "#{file['parentReference']['driveId']}/items/#{file['id']}/children" : "#{file['parentReference']['driveId']}/items/#{file['id']}"
        elsif file['id'].include?(root_site)
          "#{file['id']}/drives"
        else
          "#{file['id']}/root/children"
        end
      end

      def folder?(file)
        !file['file'].present?
      end

      def root_site
        @root_site ||= sharepoint_request("https://graph.microsoft.com/v1.0/sites/root?select=siteCollection")['siteCollection']['hostname']
      end

      def sites
        @sites ||= sharepoint_request("https://graph.microsoft.com/v1.0/sites?search=")['value']
      end

      def drives
        @drives = sharepoint_request("https://graph.microsoft.com/v1.0/me/drives")['value']
      end

      def items_by_id(id)
        if id.include?(root_site)
          item = sharepoint_request("https://graph.microsoft.com/v1.0/sites/#{id}")
        else
          item = sharepoint_request("https://graph.microsoft.com/v1.0/me/drives/#{id}")
        end
        item['value'].present? ? item['value'] : item
      end

      def download_url(file)
        file['@microsoft.graph.downloadUrl']
      end
    end
  end
end