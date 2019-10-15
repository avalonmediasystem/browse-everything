# frozen_string_literal: true

require 'rails/generators'
require 'fileutils'

class BrowseEverything::InstallGenerator < Rails::Generators::Base
  desc 'This generator installs the browse everything configuration into your application'

  source_root File.expand_path('templates', __dir__)

  def inject_config
    generate 'browse_everything:config'
  end

  def copy_migrations
    # This needs to be run later?
    rake "browse_everything:install:migrations"
  end

  def install_webpacker
    rake "webpacker:install"
  end

  # This should be removed with --skip-turbolinks, and that is passed in
  # .engine_cart.yml
  def remove_turbolinks
    gsub_file('Gemfile', /gem 'turbolinks'.*$/, '')
    # This is specific to Rails 5.2.z releases
    if Rails.version =~ /^5\./
      gsub_file('app/assets/javascripts/application.js', /\/\/= require turbolinks.*$/, '')
    elsif File.exists?(Rails.root.join('app', 'assets', 'javascripts', 'application.js'))
      # This is specific to Rails 6.y.z releases
      gsub_file('app/assets/javascripts/application.js', /require\("turbolinks".*$/, '')
    end
  end

  def install_active_storage
    rake "active_storage:install"
  end

  def install_rswag
    # This is needed for a bug, as rswag will not install for the dependent app.
    # unless it is explicitly required here
    insert_into_file 'config/application.rb', after: 'require "rails/test_unit/railtie"' do
      "\nrequire 'rswag'"
    end

    generate 'rswag:install'
  end

  # Things get more complicated here with RSpec
  # Need to install rspec, rspec-rails
  def install_rspec
    exec 'rspec --init'
    insert_into_file 'spec/spec_helper.rb', before: 'RSpec.configure do |config|' do
      "\nrequire 'rspec'\n require 'rspec-rails'"
    end
    gsub_file 'spec/swagger_helper.rb',
      "config.swagger_root = Rails.root.join('swagger').to_s",
      "rails_root_path = Pathname.new(File.dirname(__FILE__))\nconfig.swagger_root = rails_root_path.join('..', 'swagger').to_s"
  end

  def install_swagger_api_spec
    FileUtils.mkdir_p 'swagger/v1'
    copy_file '../swagger/v1/swagger.json', 'swagger/v1/swagger.json'
  end

  def install_swagger_tests
    FileUtils.mkdir_p 'spec/integration'
    Dir.glob("../spec/integration/*_spec.rb").each do |test_file_path|
      segments = test_file_path.split('/')
      target_segments = segments[1..]
      target_path = target_segments.join('/')
      copy_file test_file_path, target_path
    end
  end
end
