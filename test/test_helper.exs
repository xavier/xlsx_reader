ExUnit.start()

defmodule TestFixtures do
  def read!(relative_path) do
    relative_path
    |> path
    |> File.read!()
  end

  def path(relative_path) do
    Path.join([__DIR__, "fixtures", relative_path])
  end
end
