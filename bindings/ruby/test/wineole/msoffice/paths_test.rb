require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/paths'

class PathsTest < Minitest::Test
  def loopback_client
    Struct.new(:loopback?).new(true)
  end

  def remote_client
    Struct.new(:loopback?).new(false)
  end

  def winepath?
    return @winepath if defined?(@winepath)

    @winepath = system('which winepath > /dev/null 2>&1')
  end

  # --- convertible? -------------------------------------------------------

  def test_paths_are_not_converted_when_the_client_runs_on_windows
    refute WineOLE::MSOffice::Paths.convertible?(client: loopback_client, windows: true),
      'a Windows client already has Windows paths, and there is no winepath'
  end

  def test_paths_are_not_converted_for_a_remote_bridge
    refute WineOLE::MSOffice::Paths.convertible?(client: remote_client, windows: false),
      "the client's own path means nothing on another machine"
  end

  def test_paths_are_converted_for_a_local_bridge
    assert WineOLE::MSOffice::Paths.convertible?(client: loopback_client, windows: false)
  end

  # --- to_wine: already-Windows paths are left alone -----------------------

  def test_an_already_windows_path_is_left_alone
    assert_equal 'Z:\\home\\user\\out.xls',
      WineOLE::MSOffice::Paths.to_wine('Z:\\home\\user\\out.xls')
    assert_equal 'C:\\Temp\\x.xls', WineOLE::MSOffice::Paths.to_wine('C:\\Temp\\x.xls')
    assert_equal '\\\\server\\share\\x.xls',
      WineOLE::MSOffice::Paths.to_wine('\\\\server\\share\\x.xls')
  end

  # --- to_local: already-local (non-Windows-shaped) paths are left alone ---

  def test_a_non_windows_shaped_path_is_left_alone_by_to_local
    assert_equal '/home/user/out.xls', WineOLE::MSOffice::Paths.to_local('/home/user/out.xls')
  end

  def test_an_empty_path_is_returned_unchanged_by_to_local
    # An unsaved Workbook's Path is "" -- to_local must not shell out for it.
    assert_equal '', WineOLE::MSOffice::Paths.to_local('')
  end

  # --- round trip via the real winepath, skipped if it is not installed ---

  def test_a_linux_path_round_trips_through_winepath
    skip 'winepath is not on PATH' unless winepath?

    # A drive root (e.g. the Z: mapping of a whole prefix) round-trips with
    # a trailing slash added by winepath itself, so exercise a subdirectory
    # instead -- the round trip this method promises.
    path = File.expand_path('..', __dir__)
    wine_path = WineOLE::MSOffice::Paths.to_wine(path)
    assert_equal path, WineOLE::MSOffice::Paths.to_local(wine_path)
  end
end
